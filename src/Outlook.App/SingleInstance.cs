using System.Security.Cryptography;
using System.Text;

namespace Outlook.App;

/// <summary>
/// One running copy per data folder. Starting the app again (Start menu, desktop shortcut) while it
/// sits in the notification area brings the existing window forward instead of opening a second one
/// on the same cache. A different OUTLOOKCLONE_DATA folder is a different instance.
/// </summary>
public sealed class SingleInstance : IDisposable
{
    private readonly Mutex _mutex;
    private readonly EventWaitHandle _showSignal;
    private readonly CancellationTokenSource _stop = new();

    private SingleInstance(Mutex mutex, EventWaitHandle showSignal)
    {
        _mutex = mutex;
        _showSignal = showSignal;
    }

    /// <summary>Raised (on a background thread) when another launch asks this one to show itself.</summary>
    public event Action? ShowRequested;

    /// <summary>The instance for this data folder, or null after asking the one already running to show itself.</summary>
    public static SingleInstance? TryStart(string dataDir)
    {
        var id = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(dataDir).ToUpperInvariant())))[..16];
        var mutex = new Mutex(true, $@"Local\OutlookClone-{id}", out var first);
        var showSignal = new EventWaitHandle(false, EventResetMode.AutoReset, $@"Local\OutlookClone-{id}-show");
        if (!first)
        {
            showSignal.Set();
            showSignal.Dispose();
            mutex.Dispose();
            return null;
        }
        var instance = new SingleInstance(mutex, showSignal);
        new Thread(instance.Listen) { IsBackground = true, Name = "Single instance" }.Start();
        return instance;
    }

    private void Listen()
    {
        var handles = new[] { _showSignal, _stop.Token.WaitHandle };
        while (WaitHandle.WaitAny(handles) == 0)
            ShowRequested?.Invoke();
    }

    public void Dispose()
    {
        _stop.Cancel();
        _mutex.ReleaseMutex();
        _mutex.Dispose();
        _showSignal.Dispose();
    }
}
