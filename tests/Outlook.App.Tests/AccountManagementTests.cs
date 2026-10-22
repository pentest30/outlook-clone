using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Models;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.App.Tests;

public sealed class AccountManagementTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), $"outlook-accounts-{Guid.NewGuid():N}");
    private readonly MailStore _store;
    private readonly InMemoryCredentialStore _credentials = new();
    private readonly AppSettings _settings;

    public AccountManagementTests()
    {
        Directory.CreateDirectory(_dir);
        _store = new MailStore(() => MailDbContext.Open(Path.Combine(_dir, "cache.db")));
        _store.Migrate();
        _settings = AppSettings.Load(Path.Combine(_dir, "settings.json"));
    }

    private async Task<Account> Add(string email)
    {
        var account = await _store.AddAccountAsync(new Account { Email = email, DisplayName = email, ImapHost = "imap.x.test", SmtpHost = "smtp.x.test" });
        _credentials.Save(account.Id, "pw");
        return account;
    }

    private MainWindowViewModel Start()
    {
        var vm = new MainWindowViewModel(new FakeMail(), _store, _credentials, _settings, new FakeCalendar(),
            new ServerDiscovery(new HttpClient(), (_, _) => Task.FromResult<IReadOnlyList<string>>([])))
        {
            Confirm = _ => Task.FromResult(true),
        };
        Settle(() => vm.Accounts.Count > 0 || vm.IsAccountSetupVisible);
        return vm;
    }

    private static void Settle(Func<bool> until)
    {
        var deadline = DateTime.UtcNow.AddSeconds(5);
        while (!until() && DateTime.UtcNow < deadline)
        {
            Dispatcher.UIThread.RunJobs();
            Thread.Sleep(10);
        }
        Dispatcher.UIThread.RunJobs();
    }

    [AvaloniaFact]
    public async Task Opens_the_remembered_account_and_switches()
    {
        await Add("first@x.test");
        var second = await Add("second@x.test");
        _settings.ActiveAccountId = second.Id;

        var vm = Start();
        Assert.Equal("second@x.test", vm.AccountEmail);
        Assert.Equal([false, true], vm.Accounts.Select(a => a.IsActive));

        await vm.SwitchToCommand.ExecuteAsync(vm.Accounts[0]);

        Assert.Equal("first@x.test", vm.AccountEmail);
        Assert.Equal([true, false], vm.Accounts.Select(a => a.IsActive));
        Assert.Equal(vm.Accounts[0].Account.Id, _settings.ActiveAccountId);   // remembered for next time
    }

    [AvaloniaFact]
    public async Task Removing_the_open_account_opens_another_and_forgets_its_password()
    {
        var first = await Add("first@x.test");
        await Add("second@x.test");
        var vm = Start();
        Assert.Equal("first@x.test", vm.AccountEmail);

        await vm.RemoveAccountCommand.ExecuteAsync(vm.Accounts[0]);

        Assert.Equal("second@x.test", vm.AccountEmail);
        Assert.Equal(["second@x.test"], vm.Accounts.Select(a => a.Email));
        Assert.Null(_credentials.Load(first.Id));
    }

    [AvaloniaFact]
    public async Task Removing_the_last_account_brings_back_the_first_run_form()
    {
        await Add("only@x.test");
        var vm = Start();

        await vm.RemoveAccountCommand.ExecuteAsync(vm.Accounts[0]);

        Assert.Empty(vm.Accounts);
        Assert.True(vm.IsAccountSetupVisible);
        Assert.NotNull(vm.AccountSetup);
        Assert.False(vm.AccountSetup!.CanCancel);   // nothing to go back to
        Assert.Equal("", vm.AccountEmail);
    }

    [AvaloniaFact]
    public async Task Accounts_page_renders()
    {
        await Add("dev@localhost");
        await Add("sam.taylor@example.com");
        var vm = Start();
        var window = new SettingsWindow { DataContext = new SettingsViewModel(_settings, vm.AccountEmail, vm) };
        window.Show();
        Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        Assert.Equal(0, window.FindControl<TabControl>("Tabs")!.SelectedIndex);
        window.CaptureRenderedFrame()?.Save(Path.Combine(Path.GetTempPath(), "outlook-ui-shots", "accounts.png"));
        window.Close();
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        try { Directory.Delete(_dir, true); } catch (IOException) { }
    }
}
