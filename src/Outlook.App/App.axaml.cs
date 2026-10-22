using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Avalonia.Platform;
using Avalonia.Threading;
using Microsoft.Extensions.DependencyInjection;
using Outlook.App.ViewModels;
using Outlook.Calendar;
using Outlook.App.Views;
using Outlook.Core.Services;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.App;

public partial class App : Application
{
    private ServiceProvider? _services;

    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    /// <summary>"Light", "Dark", or anything else to follow the operating system.</summary>
    public static void ApplyTheme(string theme)
    {
        if (Current is null) return;
        Current.RequestedThemeVariant = theme switch
        {
            "Light" => Avalonia.Styling.ThemeVariant.Light,
            "Dark" => Avalonia.Styling.ThemeVariant.Dark,
            _ => Avalonia.Styling.ThemeVariant.Default,
        };
    }

    public override void OnFrameworkInitializationCompleted()
    {
        _services = ConfigureServices();

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            ApplyTheme(_services.GetRequiredService<AppSettings>().Theme);
            var vm = _services.GetRequiredService<MainWindowViewModel>();
            var window = new MainWindow { DataContext = vm };
            desktop.MainWindow = window;
            CreateTrayIcon(desktop, window, vm);

            // A bug in one handler should not take down the app and any unsaved compose windows with it.
            // Log it, tell the user, and keep running.
            Dispatcher.UIThread.UnhandledException += (_, e) =>
            {
                LogError(e.Exception);
                vm.StatusText = $"Something went wrong: {e.Exception.Message} (details in errors.log)";
                e.Handled = true;
            };
            // Run on the thread pool: blocking the UI thread on async disposal would deadlock.
            desktop.ShutdownRequested += (_, _) => Task.Run(() => _services.DisposeAsync().AsTask()).Wait(TimeSpan.FromSeconds(3));
        }

        base.OnFrameworkInitializationCompleted();
    }

    /// <summary>Notification-area icon: unread count in the tooltip; open, new mail or exit from its menu.</summary>
    private void CreateTrayIcon(IClassicDesktopStyleApplicationLifetime desktop, MainWindow window, MainWindowViewModel vm)
    {
        var open = new NativeMenuItem("Open Outlook Clone");
        open.Click += (_, _) => window.BringToFront();
        var newMail = new NativeMenuItem("New mail");
        newMail.Click += (_, _) =>
        {
            window.BringToFront();
            vm.NewMailCommand.Execute(null);
        };
        var exit = new NativeMenuItem("Exit");
        exit.Click += (_, _) => desktop.Shutdown();

        var tray = new TrayIcon
        {
            Icon = new WindowIcon(AssetLoader.Open(new Uri("avares://OutlookClone/Assets/app.ico"))),
            ToolTipText = "Outlook Clone",
            Menu = new NativeMenu { Items = { open, newMail, new NativeMenuItemSeparator(), exit } },
        };
        tray.Clicked += (_, _) => window.BringToFront();
        vm.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(MainWindowViewModel.InboxUnreadCount))
                tray.ToolTipText = vm.InboxUnreadCount == 0 ? "Outlook Clone" : $"Outlook Clone · {vm.InboxUnreadCount} unread";
        };
        TrayIcon.SetIcons(this, [tray]);
        desktop.ShutdownRequested += (_, _) => tray.Dispose();
    }

    // Everything local lives under %LOCALAPPDATA%\OutlookClone, like Outlook's own data folder.
    // OUTLOOKCLONE_DATA points it elsewhere, e.g. to keep a test profile apart from a real one.
    private static readonly string DataDir =
        Environment.GetEnvironmentVariable("OUTLOOKCLONE_DATA") is { Length: > 0 } custom
            ? custom
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookClone");

    private static void LogError(Exception ex)
    {
        try
        {
            File.AppendAllText(Path.Combine(DataDir, "errors.log"), $"{DateTimeOffset.Now:O} {ex}{Environment.NewLine}{Environment.NewLine}");
        }
        catch (IOException)
        {
            // Nowhere left to report it.
        }
    }

    private static ServiceProvider ConfigureServices()
    {
        var dataDir = DataDir;
        Directory.CreateDirectory(dataDir);
        var dbPath = Path.Combine(dataDir, "mailcache.db");

        var services = new ServiceCollection();
        Func<MailDbContext> createContext = () => MailDbContext.Open(dbPath);

        services.AddSingleton(createContext);
        services.AddSingleton(_ =>
        {
            var store = new MailStore(createContext);
            store.Migrate();
            return store;
        });
        services.AddSingleton<IMailStore>(sp => sp.GetRequiredService<MailStore>());
        services.AddSingleton<ICredentialStore>(_ => OperatingSystem.IsWindows()
            ? new DpapiCredentialStore(Path.Combine(dataDir, "credentials"))
            : new InMemoryCredentialStore());
        services.AddSingleton<ImapMailService>();
        services.AddSingleton<IMailService>(sp => sp.GetRequiredService<ImapMailService>());
        services.AddSingleton<ICalendarService, CalendarService>();
        services.AddSingleton(_ => AppSettings.Load(Path.Combine(dataDir, "settings.json")));
        services.AddSingleton<MainWindowViewModel>();

        return services.BuildServiceProvider();
    }
}
