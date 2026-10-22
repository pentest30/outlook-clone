using System.Net.Sockets;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Security;
using Outlook.Core.Models;

namespace Outlook.Mail;

/// <summary>
/// One logged-in IMAP session per account. ImapClient is not thread-safe, so every caller must
/// go through <see cref="UseAsync{T}"/>, which serialises access and reconnects when needed.
/// </summary>
internal sealed class ImapConnection(Account account, Func<string> getPassword) : IAsyncDisposable
{
    private ImapClient _client = new();
    private readonly SemaphoreSlim _gate = new(1, 1);

    public async Task<T> UseAsync<T>(Func<ImapClient, Task<T>> action, CancellationToken ct)
    {
        await _gate.WaitAsync(ct);
        try
        {
            await EnsureConnectedAsync(ct);
            try
            {
                return await action(_client);
            }
            catch (Exception ex) when (IsConnectionLost(ex) && !ct.IsCancellationRequested)
            {
                // Servers, proxies and NAT routers silently drop idle sessions, and we only find out
                // on the next command. Start a fresh session and try once more.
                await ResetAsync();
                await EnsureConnectedAsync(ct);
                return await action(_client);
            }
        }
        finally
        {
            _gate.Release();
        }
    }

    public Task UseAsync(Func<ImapClient, Task> action, CancellationToken ct) =>
        UseAsync(async c => { await action(c); return true; }, ct);

    private async Task EnsureConnectedAsync(CancellationToken ct)
    {
        if (!_client.IsConnected)
            await _client.ConnectAsync(account.ImapHost, account.ImapPort, ToSocketOptions(account.ImapSecurity), ct);
        if (!_client.IsAuthenticated)
            await _client.AuthenticateAsync(account.UserName, getPassword(), ct);
    }

    internal static bool IsConnectionLost(Exception ex) =>
        ex is IOException or SocketException or ServiceNotConnectedException or ImapProtocolException
        || ex.InnerException is IOException or SocketException;

    /// <summary>A client that hit a transport error may be half-broken, so replace it rather than reuse it.</summary>
    private async Task ResetAsync()
    {
        try { await _client.DisconnectAsync(false); }
        catch { /* the connection is already gone */ }
        _client.Dispose();
        _client = new ImapClient();
    }

    internal static SecureSocketOptions ToSocketOptions(TransportSecurity security) => security switch
    {
        TransportSecurity.SslOnConnect => SecureSocketOptions.SslOnConnect,
        TransportSecurity.StartTls => SecureSocketOptions.StartTls,
        _ => SecureSocketOptions.None,
    };

    public async ValueTask DisposeAsync()
    {
        if (_client.IsConnected)
        {
            try { await _client.DisconnectAsync(true); }
            catch { /* best effort on shutdown */ }
        }
        _client.Dispose();
        _gate.Dispose();
    }
}
