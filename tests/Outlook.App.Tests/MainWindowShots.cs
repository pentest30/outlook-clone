using Avalonia.Controls;
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
/// Renders the app with made-up mail and appointments, to look at layout and styling, and checks the
/// windows build with real data. PNGs go to %TEMP%/outlook-ui-shots, or to README_SHOTS when set:
///   $env:README_SHOTS = "$PWD/docs/screenshots"; dotnet test tests/Outlook.App.Tests --filter MainWindowShots
/// </summary>
public sealed class MainWindowShots : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"outlook-shots-{Guid.NewGuid():N}");
    private static readonly string Shots =
        Environment.GetEnvironmentVariable("README_SHOTS") is { Length: > 0 } custom ? custom : Path.Combine(Path.GetTempPath(), "outlook-ui-shots");

    private readonly System.Globalization.CultureInfo _culture = System.Globalization.CultureInfo.CurrentCulture;

    public MainWindowShots()
    {
        Directory.CreateDirectory(_dir);
        // The UI is in English; keep dates and day names in English too, whatever the machine's language.
        System.Globalization.CultureInfo.CurrentCulture = new System.Globalization.CultureInfo("en-GB");
    }

    private const string ReportHtml = """
        <div style="font-family: Segoe UI, sans-serif; font-size: 10.5pt; color: #242424">
          <p>Hi Sam,</p>
          <p>Here are the Q3 numbers ahead of Thursday's review. Overall we are <b>ahead of plan</b>, mostly thanks to the new
             onboarding flow.</p>
          <table style="border-collapse: collapse; margin: 8px 0">
            <tr>
              <td style="padding: 6px 14px; background: #0F6CBD; color: white">Region</td>
              <td style="padding: 6px 14px; background: #0F6CBD; color: white">Revenue</td>
              <td style="padding: 6px 14px; background: #0F6CBD; color: white">vs plan</td>
            </tr>
            <tr><td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0">North</td>
                <td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0">412 k€</td>
                <td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0; color: #107C10">+8 %</td></tr>
            <tr><td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0">South</td>
                <td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0">298 k€</td>
                <td style="padding: 6px 14px; border-bottom: 1px solid #e0e0e0; color: #107C10">+3 %</td></tr>
            <tr><td style="padding: 6px 14px">Export</td><td style="padding: 6px 14px">156 k€</td>
                <td style="padding: 6px 14px; color: #C50F1F">−2 %</td></tr>
          </table>
          <p>Two things to decide on Thursday:</p>
          <ul>
            <li>whether we extend the pilot to the South region,</li>
            <li>the hiring plan for Q4.</li>
          </ul>
          <p>The full workbook is attached.</p>
          <p>Best regards,<br>Ana Lopez<br><span style="color: #616161">Finance · Example Corp</span></p>
        </div>
        """;

    private async Task<MainWindowViewModel> SampleAsync(FakeCalendar? calendar = null)
    {
        var store = new MailStore(() => MailDbContext.Open(Path.Combine(_dir, "cache.db")));
        store.Migrate();
        var account = await store.AddAccountAsync(new Account { Email = "sam.taylor@example.com", DisplayName = "Sam Taylor" });
        await store.GetCategoriesAsync(account.Id);   // Outlook's six default categories
        await using (var db = store.CreateContext())
        {
            var inbox = new MailFolder { AccountId = account.Id, FullName = "INBOX", DisplayName = "Inbox", Special = SpecialFolder.Inbox, UnreadCount = 3 };
            db.Folders.AddRange(inbox,
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Drafts", DisplayName = "Drafts", ParentFullName = "INBOX", Special = SpecialFolder.Drafts },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Sent", DisplayName = "Sent", ParentFullName = "INBOX", Special = SpecialFolder.Sent },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Archive", DisplayName = "Archive", ParentFullName = "INBOX", Special = SpecialFolder.Archive },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Trash", DisplayName = "Trash", ParentFullName = "INBOX", Special = SpecialFolder.Trash },
                new MailFolder { AccountId = account.Id, FullName = "INBOX.Projects", DisplayName = "Projects", ParentFullName = "INBOX" });
            await db.SaveChangesAsync();
            var now = DateTimeOffset.Now;
            db.Messages.AddRange(
                Msg(inbox, 1, "Ana Lopez", "Q3 results ahead of Thursday's review", "Hi Sam, Here are the Q3 numbers ahead of Thursday's review", now.AddMinutes(-25),
                    read: false, thread: "q3", attach: true, categories: "$Cat_Blue_category", html: ReportHtml),
                Msg(inbox, 2, "Bilal Haddad", "RE: Website relaunch: final mock-ups", "Looks great! Two small comments on the pricing page", now.AddHours(-2), read: false, thread: "web"),
                Msg(inbox, 3, "Chen Wei", "RE: Website relaunch: final mock-ups", "I've uploaded the last version of the mock-ups", now.AddHours(-4), read: true, thread: "web"),
                Msg(inbox, 4, "Dana Novak", "Team offsite: venue options", "We have three options for the October offsite", now.AddHours(-6), read: true, thread: "offsite", flag: true,
                    categories: "$Cat_Green_category"),
                Msg(inbox, 5, "calendar@example.com", "Invitation: Product roadmap review", "You have been invited to Product roadmap review", now.AddDays(-1), read: false,
                    thread: "roadmap", calendar: true),
                Msg(inbox, 6, "Eve Martin", "Contract renewal: signed copy", "Please find the signed contract attached", now.AddDays(-1).AddHours(-3), read: true, thread: "contract",
                    attach: true, categories: "$Cat_Red_category"),
                Msg(inbox, 7, "Frank Osei", "Lunch next week?", "Are you free on Tuesday or Wednesday for lunch", now.AddDays(-3), read: true, thread: "lunch"),
                Msg(inbox, 8, "IT Service Desk", "Scheduled maintenance this weekend", "The VPN will be unavailable on Saturday from 22:00", now.AddDays(-8), read: true, thread: "it"));
            await db.SaveChangesAsync();
        }
        var settings = AppSettings.Load(Path.Combine(_dir, "settings.json"));
        var vm = new MainWindowViewModel(new FakeMail(), store, new InMemoryCredentialStore(), settings, calendar ?? new FakeCalendar(),
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
        bool read, string thread, bool attach = false, bool flag = false, bool calendar = false, string categories = "", string? html = null) => new()
    {
        FolderId = f.Id, Uid = uid, FromName = from.Contains('@') ? "" : from,
        FromAddress = from.Contains('@') ? from : $"{from.Split(' ')[0].ToLowerInvariant()}@example.com",
        Subject = subject, Preview = preview, Date = date, IsRead = read, ThreadId = thread, HasAttachments = attach,
        IsFlagged = flag, HasCalendar = calendar, CategoryKeywords = categories, To = "Sam Taylor <sam.taylor@example.com>",
        IsBodyDownloaded = true, BodyHtml = html ?? $"<div style=\"font-family: Segoe UI, sans-serif; font-size: 10.5pt\"><p>Hi Sam,</p><p>{preview}.</p><p>Best regards,<br>{from}</p></div>",
    };

    private static void Render(TopLevel window, string name)
    {
        for (var i = 0; i < 5; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
        Directory.CreateDirectory(Shots);
        window.CaptureRenderedFrame()?.Save(Path.Combine(Shots, name + ".png"));
    }

    private static async Task Settle()
    {
        for (var i = 0; i < 10; i++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(20);
        }
    }

    [AvaloniaFact]
    public async Task Main_window_light_and_dark()
    {
        var vm = await SampleAsync();
        var window = new MainWindow { DataContext = vm, Width = 1440, Height = 820, WindowState = WindowState.Normal };
        window.Show();

        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Light;
        vm.SelectedMessage = vm.Messages[0];
        await Settle();
        Render(window, "mail-light");

        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Dark;
        vm.SelectedMessage = vm.Messages.First(m => m.ThreadCount > 1);   // a conversation
        await Settle();
        Render(window, "mail-dark");

        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Default;
        window.ExitApp();
    }

    [AvaloniaFact]
    public async Task Calendar_work_week()
    {
        var calendar = new FakeCalendar();
        var monday = DateTime.Today.AddDays(-(((int)DateTime.Today.DayOfWeek + 6) % 7));
        void Add(int day, double from, double to, string title, string location = "", bool recurring = false, bool allDay = false)
        {
            var start = allDay ? new DateTimeOffset(DateTime.SpecifyKind(monday.AddDays(day), DateTimeKind.Unspecified), TimeSpan.Zero)
                : new DateTimeOffset(DateTime.SpecifyKind(monday.AddDays(day).AddHours(from), DateTimeKind.Local));
            var end = allDay ? start.AddDays(1) : start.AddHours(to - from);
            calendar.Occurrences.Add(new AppointmentOccurrence(new Appointment
            {
                Id = calendar.Occurrences.Count + 1, Title = title, Location = location, StartUtc = start, EndUtc = end,
                IsRecurring = recurring, IsAllDay = allDay, IsOrganizer = true,
            }, start, end));
        }
        for (var d = 0; d < 5; d++) Add(d, 9, 9.25, "Daily stand-up", "Teams room", recurring: true);
        Add(0, 10, 11.5, "Website relaunch kick-off", "Room 2.14");
        Add(1, 13, 14, "Lunch with Frank", "Café Central");
        Add(1, 14.5, 16, "Budget planning", "Room 3.02");
        Add(2, 11, 12, "1:1 with Ana");
        Add(2, 11.5, 12.5, "Vendor call", "Phone");
        Add(3, 10, 12, "Q3 results review", "Board room");
        Add(3, 15, 16, "Interview: frontend developer", "Room 1.05");
        Add(4, 0, 0, "Team offsite", allDay: true);
        Add(4, 13, 14.5, "Product roadmap review", "Room 2.14");

        var vm = await SampleAsync(calendar);
        var window = new MainWindow { DataContext = vm, Width = 1440, Height = 820, WindowState = WindowState.Normal };
        window.Show();
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Light;
        vm.ShowCalendarCommand.Execute(null);
        vm.Calendar.SetWorkWeekCommand.Execute(null);
        await vm.Calendar.ReloadAsync();
        await Settle();
        Render(window, "calendar");
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Default;
        window.ExitApp();
    }

    [AvaloniaFact]
    public async Task Open_shows_the_message_in_its_own_window()
    {
        var vm = await SampleAsync();
        var main = new MainWindow { DataContext = vm, Width = 1200, Height = 800 };
        main.Show();
        MessageWindowViewModel? opened = null;
        vm.MessageWindowRequested += m => opened = m;

        vm.SelectedMessage = vm.Messages.First(m => m.Subject.StartsWith("Q3 results"));
        await vm.OpenSelectedCommand.ExecuteAsync(null);
        Dispatcher.UIThread.RunJobs();

        Assert.NotNull(opened);
        Assert.Equal("Q3 results ahead of Thursday's review", opened.Title);
        Assert.Contains("ahead of plan", opened.Body.Html);
        Assert.True(opened.Message.IsRead);   // opening it counts as reading it

        // A window of our own on the same view model, to look at.
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Light;
        var shown = new MessageWindow { DataContext = opened, Width = 900, Height = 720 };
        shown.Show();
        await Settle();
        Render(shown, "message-window");
        shown.Close();
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Default;
        main.ExitApp();
    }

    [AvaloniaFact]
    public async Task Reply_opens_compose()
    {
        var vm = await SampleAsync();
        var main = new MainWindow { DataContext = vm, Width = 1200, Height = 800 };
        main.Show();
        ComposeViewModel? compose = null;
        vm.ComposeRequested += c => compose = c;

        vm.SelectedMessage = vm.Messages.First(m => m.Subject.StartsWith("Lunch"));
        await vm.ReplyCommand.ExecuteAsync(null);
        Assert.NotNull(compose);

        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Light;
        compose.Body = "Hi Frank,\n\nWednesday works for me. 12:30 at Café Central?\n\nSam\n\n" + compose.Body.TrimStart();
        var window = new ComposeWindow { DataContext = compose, Width = 900, Height = 640 };
        window.Show();
        await Settle();
        Render(window, "compose");
        Avalonia.Application.Current!.RequestedThemeVariant = ThemeVariant.Default;
        main.ExitApp();
    }

    [AvaloniaFact]
    public async Task Closing_hides_to_the_tray_and_exit_quits()
    {
        var vm = await SampleAsync();
        var window = new MainWindow { DataContext = vm };
        window.Show();
        var closed = false;
        window.Closed += (_, _) => closed = true;

        window.Close();
        Dispatcher.UIThread.RunJobs();
        Assert.False(closed);
        Assert.False(window.IsVisible);

        window.BringToFront();
        Assert.True(window.IsVisible);

        window.ExitApp();
        Dispatcher.UIThread.RunJobs();
        Assert.True(closed);
    }

    public void Dispose()
    {
        System.Globalization.CultureInfo.CurrentCulture = _culture;
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { Directory.Delete(_dir, true); } catch (IOException) { }
    }
}
