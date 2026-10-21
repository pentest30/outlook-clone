using CommunityToolkit.Mvvm.ComponentModel;

namespace Outlook.App.ViewModels;

/// <summary>A message waiting out its undo-send delay.</summary>
public partial class PendingSendViewModel(ComposeViewModel compose, int seconds) : ObservableObject
{
    public ComposeViewModel Compose { get; } = compose;
    public CancellationTokenSource Countdown { get; } = new();

    public string Subject => string.IsNullOrWhiteSpace(Compose.Subject) ? "(no subject)" : Compose.Subject;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Text))]
    private int _secondsLeft = seconds;

    public string Text => $"Sending “{Subject}” in {SecondsLeft} s";
}
