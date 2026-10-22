using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Compose;
using Outlook.Core.Models;

namespace Outlook.App.ViewModels;

/// <summary>A message opened in its own window (double-click or Enter in the list).</summary>
public partial class MessageWindowViewModel(MessageItemViewModel header, MessageBodyViewModel body, MainWindowViewModel host)
{
    public MessageItemViewModel Header { get; } = header;
    public MessageBodyViewModel Body { get; } = body;
    public MailMessage Message => Header.Model;
    public string Title => Header.Subject;

    /// <summary>The window closes itself after Delete.</summary>
    public event Action? CloseRequested;

    public void ComposeTo(string address) => host.ComposeTo(address);

    [RelayCommand]
    private Task ReplyAsync() => host.RespondAsync(ComposeMode.Reply, Message);

    [RelayCommand]
    private Task ReplyAllAsync() => host.RespondAsync(ComposeMode.ReplyAll, Message);

    [RelayCommand]
    private Task ForwardAsync() => host.RespondAsync(ComposeMode.Forward, Message);

    [RelayCommand]
    private async Task DeleteAsync()
    {
        if (await host.DeleteMessageAsync(Message))
            CloseRequested?.Invoke();
    }
}
