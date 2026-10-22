using System.Net;
using System.Xml.Linq;
using DnsClient;
using Outlook.Core.Models;

namespace Outlook.Mail;

/// <summary>Server settings found for an address, and where they came from.</summary>
public sealed record DiscoveredSettings(
    string ImapHost, int ImapPort, TransportSecurity ImapSecurity,
    string SmtpHost, int SmtpPort, TransportSecurity SmtpSecurity,
    string UserName, string Provider, bool NeedsOAuth);

/// <summary>
/// Finds IMAP/SMTP settings from just an e-mail address, the way Thunderbird does: look the domain
/// up in Mozilla's public provider database (ISPDB); if it is not there, find who handles the
/// domain's mail (its MX records, e.g. mx1.mail.ovh.net) and look up that provider instead
/// (mail.ovh.net, then ovh.net). Company domains hosted by OVH, Gmail, Microsoft 365 and others
/// are found this way.
/// </summary>
public sealed class ServerDiscovery(HttpClient http, Func<string, CancellationToken, Task<IReadOnlyList<string>>>? mxLookup = null)
{
    public const string IspdbUrl = "https://autoconfig.thunderbird.net/v1.1/";

    private readonly Func<string, CancellationToken, Task<IReadOnlyList<string>>> _mxLookup = mxLookup ?? LookupMxAsync;

    public static ServerDiscovery CreateDefault() =>
        new(new HttpClient { Timeout = TimeSpan.FromSeconds(8) });

    public async Task<DiscoveredSettings?> DiscoverAsync(string email, CancellationToken ct = default)
    {
        var at = email.LastIndexOf('@');
        if (at <= 0 || at == email.Length - 1) return null;
        var domain = email[(at + 1)..].Trim().ToLowerInvariant();

        if (await FromIspdbAsync(domain, email, ct) is { } direct)
            return direct;

        IReadOnlyList<string> mx;
        try
        {
            mx = await _mxLookup(domain, ct);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return null;   // no DNS: the user fills the servers in by hand
        }
        foreach (var host in mx)
        {
            foreach (var candidate in ProviderDomains(host))
            {
                if (candidate == domain) continue;
                if (await FromIspdbAsync(candidate, email, ct) is { } viaMx)
                    return viaMx;
            }
        }
        return null;
    }

    /// <summary>"mx1.mail.ovh.net" gives "mail.ovh.net" then "ovh.net": the provider's own domains, most specific first.</summary>
    public static IEnumerable<string> ProviderDomains(string mxHost)
    {
        var labels = mxHost.Trim().TrimEnd('.').ToLowerInvariant().Split('.');
        for (var skip = 1; labels.Length - skip >= 2; skip++)
            yield return string.Join('.', labels[skip..]);
    }

    private async Task<DiscoveredSettings?> FromIspdbAsync(string domain, string email, CancellationToken ct)
    {
        try
        {
            using var response = await http.GetAsync(IspdbUrl + Uri.EscapeDataString(domain), ct);
            if (response.StatusCode != HttpStatusCode.OK) return null;
            return Parse(await response.Content.ReadAsStringAsync(ct), email, domain);
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException && !ct.IsCancellationRequested)
        {
            return null;
        }
    }

    /// <summary>Reads a Thunderbird autoconfig document (clientConfig v1.1).</summary>
    public static DiscoveredSettings? Parse(string xml, string email, string provider)
    {
        XDocument doc;
        try
        {
            doc = XDocument.Parse(xml);
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }
        var incoming = doc.Descendants("incomingServer").FirstOrDefault(e => (string?)e.Attribute("type") == "imap");
        var outgoing = doc.Descendants("outgoingServer").FirstOrDefault(e => (string?)e.Attribute("type") == "smtp");
        if (incoming is null || outgoing is null) return null;

        var oauth = incoming.Elements("authentication").Concat(outgoing.Elements("authentication"))
            .Any(a => a.Value.Equals("OAuth2", StringComparison.OrdinalIgnoreCase))
            && !incoming.Elements("authentication").Any(a => a.Value.StartsWith("password", StringComparison.OrdinalIgnoreCase));

        return new DiscoveredSettings(
            Value(incoming, "hostname", email), Port(incoming, 993), Security(incoming),
            Value(outgoing, "hostname", email), Port(outgoing, 587), Security(outgoing),
            Value(incoming, "username", email) is { Length: > 0 } user ? user : email,
            doc.Descendants("displayName").FirstOrDefault()?.Value ?? provider,
            oauth);
    }

    private static string Value(XElement server, string name, string email)
    {
        var at = email.IndexOf('@');
        return (server.Element(name)?.Value ?? "").Trim()
            .Replace("%EMAILADDRESS%", email)
            .Replace("%EMAILLOCALPART%", at > 0 ? email[..at] : email)
            .Replace("%EMAILDOMAIN%", at > 0 ? email[(at + 1)..] : "");
    }

    private static int Port(XElement server, int fallback) =>
        int.TryParse(server.Element("port")?.Value, out var port) ? port : fallback;

    private static TransportSecurity Security(XElement server) => server.Element("socketType")?.Value.ToUpperInvariant() switch
    {
        "SSL" => TransportSecurity.SslOnConnect,
        "STARTTLS" => TransportSecurity.StartTls,
        _ => TransportSecurity.None,
    };

    private static async Task<IReadOnlyList<string>> LookupMxAsync(string domain, CancellationToken ct)
    {
        var result = await new LookupClient().QueryAsync(domain, QueryType.MX, cancellationToken: ct);
        return result.Answers.MxRecords()
            .OrderBy(r => r.Preference)
            .Select(r => r.Exchange.Value.TrimEnd('.'))
            .ToList();
    }
}
