using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.App.Tests;

/// <summary>Accepts or refuses logins; records what was checked.</summary>
internal sealed class FakeMail : IMailService
{
    public Func<Account, string, bool> ImapAccepts { get; set; } = (_, _) => true;
    public Func<Account, string, bool> SmtpAccepts { get; set; } = (_, _) => true;
    public List<int> Forgotten { get; } = [];

    public Task VerifyAsync(Account a, string p, CancellationToken ct = default) =>
        ImapAccepts(a, p) ? Task.CompletedTask : throw new InvalidOperationException("Authentication failed.");
    public Task VerifySmtpAsync(Account a, string p, CancellationToken ct = default) =>
        SmtpAccepts(a, p) ? Task.CompletedTask : throw new InvalidOperationException("SMTP refused.");
    public Task ForgetAccountAsync(int id)
    {
        Forgotten.Add(id);
        return Task.CompletedTask;
    }

    public event EventHandler<FolderChangedEventArgs>? FolderChanged { add { } remove { } }
    public Task SyncAccountAsync(Account a, IProgress<SyncProgress>? p = null, CancellationToken ct = default) => Task.CompletedTask;
    public Task SyncFolderAsync(Account a, MailFolder f, CancellationToken ct = default) => Task.CompletedTask;
    public Task RunPushAsync(Account a, CancellationToken ct) => Task.CompletedTask;
    public Task<MailMessage> DownloadBodyAsync(MailMessage m, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SetReadAsync(MailMessage m, bool r, CancellationToken ct = default) => Task.CompletedTask;
    public Task SetFlaggedAsync(MailMessage m, bool f, CancellationToken ct = default) => Task.CompletedTask;
    public Task MarkRespondedAsync(MailMessage m, bool f, CancellationToken ct = default) => Task.CompletedTask;
    public Task SetCategoryAsync(IReadOnlyList<MailMessage> m, string k, bool a, CancellationToken ct = default) => Task.CompletedTask;
    public Task MoveAsync(IReadOnlyList<MailMessage> m, MailFolder t, CancellationToken ct = default) => Task.CompletedTask;
    public Task DeleteAsync(IReadOnlyList<MailMessage> m, CancellationToken ct = default) => Task.CompletedTask;
    public Task ArchiveAsync(IReadOnlyList<MailMessage> m, CancellationToken ct = default) => Task.CompletedTask;
    public Task<SendReport> SendAsync(Account a, OutgoingMessage m, CancellationToken ct = default) => Task.FromResult(new SendReport(true, null));
    public Task SaveDraftAsync(Account a, OutgoingMessage d, CancellationToken ct = default) => Task.CompletedTask;
    public Task DeleteDraftAsync(Account a, string id, CancellationToken ct = default) => Task.CompletedTask;
}

public sealed class AccountSetupTests : IDisposable
{
    private static readonly DiscoveredSettings Ovh = new("ssl0.ovh.net", 993, TransportSecurity.SslOnConnect,
        "ssl0.ovh.net", 465, TransportSecurity.SslOnConnect, "i.baouche@netcom-group.co", "OVH", false);

    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-setup-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;
    private readonly InMemoryCredentialStore _credentials = new();
    private readonly FakeMail _mail = new();
    private readonly List<(Account Account, bool IsNew)> _saved = [];

    public AccountSetupTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
    }

    private AccountSetupViewModel Setup(Account? existing = null, DiscoveredSettings? found = null) =>
        new(_mail, _store, _credentials, (_, _) => Task.FromResult(found), (a, isNew) =>
        {
            _saved.Add((a, isNew));
            return Task.CompletedTask;
        }, existing);

    [AvaloniaFact]
    public async Task Address_and_password_are_enough_when_servers_are_found()
    {
        var vm = Setup(found: Ovh);
        vm.Email = "i.baouche@netcom-group.co";
        vm.Password = "secret";
        await vm.DiscoverAsync();

        Assert.Equal("ssl0.ovh.net", vm.ImapHost);
        Assert.Equal(465, vm.SmtpPort);
        Assert.Equal("Found the settings for OVH (ssl0.ovh.net).", vm.DiscoveryNote);
        Assert.False(vm.ShowServerSettings);   // nothing for the user to look at

        await vm.SignInCommand.ExecuteAsync(null);

        Assert.Null(vm.Error);
        var (account, isNew) = Assert.Single(_saved);
        Assert.True(isNew);
        Assert.Equal("i.baouche@netcom-group.co", account.UserName);
        Assert.Equal("secret", _credentials.Load(account.Id));
        Assert.Equal("", vm.Password);   // not kept in memory after saving
    }

    [AvaloniaFact]
    public async Task Unknown_provider_opens_the_server_settings()
    {
        var vm = Setup(found: null);
        vm.Email = "me@small-isp.test";
        await vm.DiscoverAsync();

        Assert.True(vm.ShowServerSettings);
        Assert.StartsWith("Couldn't find the servers automatically", vm.DiscoveryNote);
    }

    [AvaloniaFact]
    public async Task A_wrong_password_or_smtp_setting_saves_nothing()
    {
        var vm = Setup(found: Ovh);
        vm.Email = "i.baouche@netcom-group.co";
        vm.Password = "wrong";
        await vm.DiscoverAsync();
        _mail.ImapAccepts = (_, p) => p == "right";

        await vm.SignInCommand.ExecuteAsync(null);
        Assert.StartsWith("Couldn't sign in to ssl0.ovh.net", vm.Error);

        vm.Password = "right";
        _mail.SmtpAccepts = (_, _) => false;
        await vm.SignInCommand.ExecuteAsync(null);
        Assert.StartsWith("Signed in, but sending through ssl0.ovh.net failed", vm.Error);

        Assert.Empty(_saved);
        Assert.Empty(await _store.GetAccountsAsync());
    }

    [AvaloniaFact]
    public async Task Editing_keeps_the_saved_password_and_reconnects()
    {
        var existing = await _store.AddAccountAsync(new Account
        {
            Email = "me@x.test", UserName = "me@x.test", ImapHost = "old.x.test", ImapPort = 993, SmtpHost = "old.x.test", SmtpPort = 465,
        });
        _credentials.Save(existing.Id, "kept");
        string? usedPassword = null;
        _mail.ImapAccepts = (_, p) => { usedPassword = p; return true; };

        var vm = Setup(existing);
        Assert.True(vm.ShowServerSettings);   // editing shows what is configured
        vm.ImapHost = "new.x.test";
        await vm.SignInCommand.ExecuteAsync(null);

        Assert.Equal("kept", usedPassword);
        Assert.Equal("new.x.test", Assert.Single(await _store.GetAccountsAsync()).ImapHost);
        Assert.Equal([existing.Id], _mail.Forgotten);
        Assert.False(Assert.Single(_saved).IsNew);
    }

    [AvaloniaFact]
    public async Task The_same_address_cannot_be_added_twice()
    {
        await _store.AddAccountAsync(new Account { Email = "me@x.test" });
        var vm = Setup(found: Ovh with { UserName = "me@x.test" });
        vm.Email = "ME@x.test";
        vm.Password = "p";
        await vm.DiscoverAsync();
        await vm.SignInCommand.ExecuteAsync(null);

        Assert.Equal("ME@x.test is already set up.", vm.Error);
    }

    [AvaloniaFact]
    public async Task Form_renders_with_detected_settings()
    {
        var vm = Setup(found: Ovh);
        var window = new AccountSetupWindow { DataContext = vm };
        window.Show();
        vm.DisplayName = "Iqbal Baouche";
        vm.Email = "i.baouche@netcom-group.co";
        await vm.DiscoverAsync();
        Dispatcher.UIThread.RunJobs();
        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "outlook-ui-shots"));
        window.CaptureRenderedFrame()?.Save(Path.Combine(Path.GetTempPath(), "outlook-ui-shots", "account-setup.png"));
        window.Close();
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
