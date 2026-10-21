using Avalonia;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Threading;
using Microsoft.Extensions.DependencyInjection;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Services;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.App;

public partial class App : Application
{
    private ServiceProvider? _services;

    public override void Initialize() => AvaloniaXamlLoader.Load(this);

    public override void OnFrameworkInitializationCompleted()
    {
        _services = ConfigureServices();

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            var vm = _services.GetRequiredService<MainWindowViewModel>();
            desktop.MainWindow = new MainWindow { DataContext = vm };

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
        services.AddSingleton(_ => AppSettings.Load(Path.Combine(dataDir, "settings.json")));
        services.AddSingleton<MainWindowViewModel>();

        return services.BuildServiceProvider();
    }
}
