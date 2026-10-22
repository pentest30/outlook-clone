using Avalonia;
using Avalonia.Controls;
using Outlook.App.ViewModels;
using Outlook.Core.Models;

namespace Outlook.App.Views;

/// <summary>
/// Shows desktop alerts (new mail, reminders) stacked upward from the bottom-right corner of the
/// screen. At most a few at once; the last one summarises the rest ("+4 more new messages").
/// </summary>
public sealed class ToastManager(Window owner, Action<MailMessage> open)
{
    public const int MaxVisible = 3;
    private const int Margin = 12;
    private const int Gap = 8;

    private readonly List<ToastWindow> _open = [];

    public IReadOnlyList<ToastWindow> Open => _open;

    public void ShowNewMail(IReadOnlyList<MailMessage> messages)
    {
        if (messages.Count == 0) return;
        var shown = messages.Take(MaxVisible).ToList();
        for (var i = 0; i < shown.Count; i++)
        {
            var m = shown[i];
            var rest = messages.Count - shown.Count;
            var heading = i == shown.Count - 1 && rest > 0 ? $"New message · +{rest} more new messages" : "New message";
            Show(new ToastContent(heading, m.SenderDisplay, string.IsNullOrWhiteSpace(m.Subject) ? "(no subject)" : m.Subject, m.Preview, m));
        }
    }

    /// <summary>Calendar reminders that just came due; clicking one opens it.</summary>
    public void ShowReminders(IReadOnlyList<ReminderItemViewModel> reminders, Action<ReminderItemViewModel> openReminder)
    {
        if (reminders.Count == 0) return;
        var shown = reminders.Take(MaxVisible).ToList();
        for (var i = 0; i < shown.Count; i++)
        {
            var r = shown[i];
            var rest = reminders.Count - shown.Count;
            var heading = i == shown.Count - 1 && rest > 0 ? $"Reminder · +{rest} more" : $"Reminder · {r.DueText}";
            Show(new ToastContent(heading, r.Title, r.When, r.Location, null, () => openReminder(r)));
        }
    }

    public void Show(ToastContent content)
    {
        // Too many already: drop the oldest so the stack never climbs the screen.
        while (_open.Count >= MaxVisible)
            _open[0].Close();

        var toast = new ToastWindow { DataContext = content };
        toast.Clicked += c =>
        {
            if (c.Activated is { } activated) activated();
            else if (c.Message is { } message) open(message);
        };
        toast.Closed += (_, _) =>
        {
            _open.Remove(toast);
            Arrange();
        };
        _open.Add(toast);
        // No owner: owned windows hide with a minimized owner, which is exactly when alerts matter.
        toast.Show();
        toast.LayoutUpdated += (_, _) => Arrange();
        Arrange();
    }

    /// <summary>Newest at the bottom, older ones pushed up.</summary>
    private void Arrange()
    {
        var screen = owner.Screens.ScreenFromWindow(owner) ?? owner.Screens.Primary;
        if (screen is null) return;
        var area = screen.WorkingArea;
        var scale = screen.Scaling;
        var bottom = area.Bottom - (int)(Margin * scale);
        for (var i = _open.Count - 1; i >= 0; i--)
        {
            var toast = _open[i];
            var size = toast.Bounds.Size == default ? new Size(toast.Width, 90) : toast.Bounds.Size;
            var width = (int)(size.Width * scale);
            var height = (int)(size.Height * scale);
            var position = new PixelPoint(area.Right - width - (int)(Margin * scale), bottom - height);
            if (toast.Position != position)
                toast.Position = position;
            bottom -= height + (int)(Gap * scale);
        }
    }
}
