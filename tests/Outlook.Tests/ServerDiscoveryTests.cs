using System.Net;
using Outlook.Core.Models;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.Tests;

public class ServerDiscoveryTests
{
    private const string OvhXml = """
        <clientConfig version="1.1">
          <emailProvider id="ovh.net">
            <displayName>OVH</displayName>
            <incomingServer type="pop3"><hostname>ssl0.ovh.net</hostname><port>995</port><socketType>SSL</socketType><username>%EMAILADDRESS%</username></incomingServer>
            <incomingServer type="imap"><hostname>ssl0.ovh.net</hostname><port>993</port><socketType>SSL</socketType><username>%EMAILADDRESS%</username><authentication>password-cleartext</authentication></incomingServer>
            <outgoingServer type="smtp"><hostname>ssl0.ovh.net</hostname><port>465</port><socketType>SSL</socketType><username>%EMAILADDRESS%</username></outgoingServer>
          </emailProvider>
        </clientConfig>
        """;

    /// <summary>Serves only the given ISPDB entries; everything else is 404.</summary>
    private sealed class FakeIspdb(Dictionary<string, string> entries) : HttpMessageHandler
    {
        public List<string> Asked { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            var domain = request.RequestUri!.AbsolutePath.Split('/').Last();
            Asked.Add(domain);
            return Task.FromResult(entries.TryGetValue(domain, out var xml)
                ? new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(xml) }
                : new HttpResponseMessage(HttpStatusCode.NotFound));
        }
    }

    [Fact]
    public void Reads_imap_and_smtp_and_skips_pop()
    {
        var s = ServerDiscovery.Parse(OvhXml, "i.baouche@netcom-group.co", "ovh.net")!;

        Assert.Equal(("ssl0.ovh.net", 993, TransportSecurity.SslOnConnect), (s.ImapHost, s.ImapPort, s.ImapSecurity));
        Assert.Equal(("ssl0.ovh.net", 465, TransportSecurity.SslOnConnect), (s.SmtpHost, s.SmtpPort, s.SmtpSecurity));
        Assert.Equal("i.baouche@netcom-group.co", s.UserName);
        Assert.Equal("OVH", s.Provider);
        Assert.False(s.NeedsOAuth);
    }

    [Fact]
    public void Placeholders_and_starttls_are_understood()
    {
        var xml = """
            <clientConfig><emailProvider>
              <incomingServer type="imap"><hostname>imap.%EMAILDOMAIN%</hostname><port>143</port><socketType>STARTTLS</socketType><username>%EMAILLOCALPART%</username></incomingServer>
              <outgoingServer type="smtp"><hostname>smtp.x.test</hostname><port>587</port><socketType>STARTTLS</socketType></outgoingServer>
            </emailProvider></clientConfig>
            """;
        var s = ServerDiscovery.Parse(xml, "bob@x.test", "x.test")!;

        Assert.Equal("imap.x.test", s.ImapHost);
        Assert.Equal(TransportSecurity.StartTls, s.ImapSecurity);
        Assert.Equal("bob", s.UserName);
    }

    [Fact]
    public void Oauth_only_providers_are_flagged()
    {
        var xml = """
            <clientConfig><emailProvider>
              <incomingServer type="imap"><hostname>outlook.office365.com</hostname><port>993</port><socketType>SSL</socketType><authentication>OAuth2</authentication></incomingServer>
              <outgoingServer type="smtp"><hostname>smtp.office365.com</hostname><port>587</port><socketType>STARTTLS</socketType><authentication>OAuth2</authentication></outgoingServer>
            </emailProvider></clientConfig>
            """;
        Assert.True(ServerDiscovery.Parse(xml, "a@contoso.com", "contoso.com")!.NeedsOAuth);
    }

    [Theory]
    [InlineData("mx1.mail.ovh.net", new[] { "mail.ovh.net", "ovh.net" })]
    [InlineData("aspmx.l.google.com.", new[] { "l.google.com", "google.com" })]
    [InlineData("mail.example.com", new[] { "example.com" })]
    public void Provider_domains_come_from_the_mail_server_name(string mx, string[] expected)
    {
        Assert.Equal(expected, ServerDiscovery.ProviderDomains(mx));
    }

    [Fact]
    public async Task Company_domain_is_found_through_its_mail_servers()
    {
        var ispdb = new FakeIspdb(new() { ["ovh.net"] = OvhXml });
        var discovery = new ServerDiscovery(new HttpClient(ispdb),
            (domain, _) => Task.FromResult<IReadOnlyList<string>>(domain == "netcom-group.co" ? ["mx1.mail.ovh.net", "mx2.mail.ovh.net"] : []));

        var s = await discovery.DiscoverAsync("i.baouche@netcom-group.co");

        Assert.Equal("ssl0.ovh.net", s?.ImapHost);
        Assert.Equal(["netcom-group.co", "mail.ovh.net", "ovh.net"], ispdb.Asked);   // the domain first, then the provider
    }

    [Fact]
    public async Task Unknown_domains_and_dns_failures_return_nothing()
    {
        var none = new ServerDiscovery(new HttpClient(new FakeIspdb([])), (_, _) => Task.FromResult<IReadOnlyList<string>>([]));
        Assert.Null(await none.DiscoverAsync("a@nowhere.test"));
        Assert.Null(await none.DiscoverAsync("not-an-address"));

        var broken = new ServerDiscovery(new HttpClient(new FakeIspdb([])), (_, _) => throw new InvalidOperationException("no DNS"));
        Assert.Null(await broken.DiscoverAsync("a@nowhere.test"));
    }
}

public sealed class AccountStoreTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-acct-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;

    public AccountStoreTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
    }

    [Fact]
    public async Task Removing_an_account_deletes_everything_cached_for_it()
    {
        var keep = await _store.AddAccountAsync(new Account { Email = "keep@x.test" });
        var gone = await _store.AddAccountAsync(new Account { Email = "gone@x.test" });
        await using (var db = _store.CreateContext())
        {
            var folder = new MailFolder { AccountId = gone.Id, FullName = "INBOX" };
            db.Folders.Add(folder);
            await db.SaveChangesAsync();
            db.Messages.Add(new MailMessage { FolderId = folder.Id, Uid = 1, Subject = "x", ThreadId = "t" });
            db.Appointments.Add(new Appointment { AccountId = gone.Id, Uid = "u", Title = "Meeting" });
            db.KnownAddresses.Add(new KnownAddress { AccountId = gone.Id, Address = "ana@x.test" });
            await db.SaveChangesAsync();
        }
        await _store.GetCategoriesAsync(gone.Id);

        await _store.RemoveAccountAsync(gone.Id);

        await using var check = _store.CreateContext();
        Assert.Equal([keep.Id], check.Accounts.Select(a => a.Id));
        Assert.Empty(check.Folders);
        Assert.Empty(check.Messages);
        Assert.Empty(check.Appointments);
        Assert.Empty(check.KnownAddresses);
        Assert.Empty(check.Categories);
    }

    [Fact]
    public async Task Updating_an_account_changes_its_servers()
    {
        var account = await _store.AddAccountAsync(new Account { Email = "me@x.test", ImapHost = "old" });
        account.ImapHost = "ssl0.ovh.net";
        account.SmtpPort = 465;
        await _store.UpdateAccountAsync(account);

        var saved = Assert.Single(await _store.GetAccountsAsync());
        Assert.Equal(("ssl0.ovh.net", 465), (saved.ImapHost, saved.SmtpPort));
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
