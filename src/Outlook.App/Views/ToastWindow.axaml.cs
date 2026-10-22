using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Outlook.Core.Models;

namespace Outlook.App.Views;

/// <summary>What one alert shows.</summary>
public sealed record ToastContent(string Heading, string Sender, string Subject, string Preview, MailMessage? Message)
{
    public bool HasPreview => Preview.Length > 0;
}

public partial class ToastWindow : Window
{
    /// <summary>How long an alert stays up, as in Outlook's default desktop alert.</summary>
    public static readonly TimeSpan Lifetime = TimeSpan.FromSeconds(7);

    private readonly DispatcherTimer _timer = new() { Interval = Lifetime };

    public ToastWindow()
    {
        InitializeComponent();
        _timer.Tick += (_, _) => Close();
        // Hovering keeps the alert up so it can be read; leaving restarts the countdown.
        PointerEntered += (_, _) => _timer.Stop();
        PointerExited += (_, _) => _timer.Start();
        CloseButton.Click += (_, e) =>
        {
            e.Handled = true;
            Close();
        };
    }

    /// <summary>The alert body was clicked (not the dismiss button).</summary>
    public event Action<ToastContent>? Clicked;

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        _timer.Start();
    }

    protected override void OnClosed(EventArgs e)
    {
        _timer.Stop();
        base.OnClosed(e);
    }

    protected override void OnPointerReleased(PointerReleasedEventArgs e)
    {
        base.OnPointerReleased(e);
        if (e.Handled || e.InitialPressMouseButton != MouseButton.Left || DataContext is not ToastContent content) return;
        Clicked?.Invoke(content);
        Close();
    }

    /// <summary>Lets tests click the alert without a pointer.</summary>
    internal void Click()
    {
        if (DataContext is ToastContent content) Clicked?.Invoke(content);
        Close();
    }
}
