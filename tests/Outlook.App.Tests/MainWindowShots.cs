using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Styling;
using Avalonia.Threading;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Models;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.App.Tests;

/// <summary>
/// Renders the main window with sample mail in light and dark, to look at layout and styling
/// (writes PNGs to %TEMP%/outlook-ui-shots). Also checks the window builds with real data.
/// </summary>
public sealed class MainWindowShots : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"outlook-shots-{Guid.NewGuid():N}");
    private static readonly string Shots = Path.Combine(Path.GetTempPath(), "outlook-ui-shots");

    public MainWindowShots() => Directory.CreateDirectory(_dir);

    private async Task<MainWindowViewModel> SampleAsync()
    {
        var store = new MailStore(() => MailDbContext.Open(Path.Combine(_dir, "cache.db")));
        store.Migrate();
        var account = await store.AddAccountAsync(new Account { Email = "i.baouche@netcom-group.co", DisplayName = "Iqbal Baouche" });
        await using (var db = store.CreateContext())
        {
            var inbox = new MailFolder { AccountId = account.Id, FullName = "INBOX", DisplayName = "Inbox", Special = SpecialFolder.Inbox, UnreadCount = 2 };
            var sent = new MailFolder { AccountId = account.Id, FullName = "INBOX.Sent", DisplayName = "Sent", ParentFullName = "INBOX", Special = SpecialFolder.Sent };
            db.Folders.AddRange(inbox, sent,
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Drafts", DisplayName = "Drafts", ParentFullName = "INBOX", Special = SpecialFolder.Drafts },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Trash", DisplayName = "Trash", ParentFullName = "INBOX", Special = SpecialFolder.Trash },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Projects", DisplayName = "Projects", ParentFullName = "INBOX" });
            await db.SaveChangesAsync();
            var now = DateTimeOffset.Now;
            db.Messages.AddRange(
                Msg(inbox, 1, "Fredrik Salén", "Sv: [Full Nexera] - Voice usage identification in TAP3 CDR", "Hi Paul, I've created an issue internally ORDER-110 to estimate", now.AddMinutes(-40), read: false, thread: "nexera", attach: true),
                Msg(inbox, 2, "Paul BUSUIOC", "RE: [Full Nexera] - Voice usage identification in TAP3 CDR", "Hello Fredrik, after an internal discussion", now.AddHours(-3), read: true, thread: "nexera"),
                Msg(inbox, 3, "messenger@webex.com", "Rappel de réunion Webex : Ikbal BAOUCHE En réunion", "Bonjour Ikbal BAOUCHE, Vous êtes programmé(e) pour organiser", now.AddHours(-5), read: true, thread: "w1", calendar: true),
                Msg(inbox, 4, "Anthropic", "Votre lien sécurisé vers Claude est ici", "Connectez-vous à Claude Cliquez sur le bouton ci-dessous", now.AddDays(-1), read: false, thread: "a1", flag: true),
                Msg(inbox, 5, "Paul BUSUIOC", "[FULL MVNO - Tache #26179] [Full] - Vérification", "La demande #26179 a été mise à jour par Paul BUSUIOC", now.AddDays(-3), read: true, thread: "t2"),
                Msg(inbox, 6, "Jérome ASSUIED", "RE: CR 15/10 Call suivi Réclamations / Facturation", "A tous, 1. Actions en cours * CDC en cours pour gérer", now.AddDays(-8), read: true, thread: "t3"));
            await db.SaveChangesAsync();
        }
        var settings = AppSettings.Load(Path.Combine(_dir, "settings.json"));
        var vm = new MainWindowViewModel(new FakeMail(), store, new InMemoryCredentialStore(), settings, new FakeCalendar(),
            new ServerDiscovery(new HttpClient(), (_, _) => Task.FromResult<IReadOnlyList<string>>([])));
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (vm.Messages.Count == 0 && DateTime.UtcNow < deadline)
        {
            Dispatcher.UIThread.RunJobs();
            Thread.Sleep(10);
        }
        return vm;
    }

    private static MailMessage Msg(MailFolder f, uint uid, string from, string subject, string preview, DateTimeOffset date,
        bool read, string thread, bool attach = false, bool flag = false, bool calendar = false) => new()
    {
        FolderId = f.Id, Uid = uid, FromName = from.Contains('@') ? "" : from, FromAddress = from.Contains('@') ? from : $"{uid}@x.test",
        Subject = subject, Preview = preview, Date = date, IsRead = read, ThreadId = thread, HasAttachments = attach,
        IsFlagged = flag, HasCalendar = calendar, To = "Iqbal Baouche <i.baouche@netcom-group.co>",
    };

    private static void Render(MainWindow window, string name)
    {
        for (var i = 0; i < 5; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
        Directory.CreateDirectory(Shots);
        window.CaptureRenderedFrame()?.Save(Path.Combine(Shots, name + ".png"));
    }

    [AvaloniaFact]
    public async Task Main_window_light_and_dark()
    {
        var vm = await SampleAsync();
        var window = new MainWindow { DataContext = vm, Width = 1500, Height = 820, WindowState = Avalonia.Controls.WindowState.Normal };
        window.Show();

        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Light;
        Render(window, "main-light");
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Dark;
        Render(window, "main-dark");

        vm.SelectedMessage = vm.Messages[1];
        Render(window, "main-dark-selected");
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Default;
        window.Close();
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { Directory.Delete(_dir, true); } catch (IOException) { }
    }
}
