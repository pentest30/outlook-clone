using System.Net;
using System.Net.Sockets;

namespace Outlook.Tests;

/// <summary>
/// Forwards TCP connections to a real server and can cut them all, the way a NAT router or
/// a server timeout kills an idle IMAP session.
/// </summary>
public sealed class DroppingProxy : IAsyncDisposable
{
    private readonly TcpListener _listener = new(IPAddress.Loopback, 0);
    private readonly List<TcpClient> _connections = [];
    private readonly CancellationTokenSource _cts = new();
    private readonly Task _acceptLoop;

    public DroppingProxy(string host, int port)
    {
        _listener.Start();
        _acceptLoop = AcceptAsync(host, port);
    }

    public int Port => ((IPEndPoint)_listener.LocalEndpoint).Port;

    public void DropAll()
    {
        lock (_connections)
        {
            foreach (var c in _connections)
                c.Close();
            _connections.Clear();
        }
    }

    private async Task AcceptAsync(string host, int port)
    {
        while (!_cts.IsCancellationRequested)
        {
            TcpClient inbound;
            try { inbound = await _listener.AcceptTcpClientAsync(_cts.Token); }
            catch (OperationCanceledException) { return; }

            var outbound = new TcpClient();
            await outbound.ConnectAsync(host, port);
            lock (_connections)
                _connections.AddRange([inbound, outbound]);
            _ = PipeAsync(inbound, outbound);
            _ = PipeAsync(outbound, inbound);
        }
    }

    private static async Task PipeAsync(TcpClient from, TcpClient to)
    {
        try { await from.GetStream().CopyToAsync(to.GetStream()); }
        catch (Exception) { /* dropped on purpose */ }
        to.Close();
    }

    public async ValueTask DisposeAsync()
    {
        await _cts.CancelAsync();
        _listener.Stop();
        DropAll();
        try { await _acceptLoop; } catch { /* stopping */ }
    }
}
