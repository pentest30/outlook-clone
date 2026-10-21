using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

/// <summary>
/// One message in the reading pane's conversation stack. Collapsed cards show a one-line preview;
/// expanding downloads and renders the body, as in Outlook's conversation view.
/// </summary>
public partial class ConversationCardViewModel : ObservableObject
{
    public ConversationCardViewModel(MailMessage message, string? folderLabel, IMailService mail, IMailStore store)
    {
        Message = message;
        FolderLabel = folderLabel;
        Body = new MessageBodyViewModel(mail, store);
    }

    public MailMessage Message { get; }
    public MessageBodyViewModel Body { get; }

    /// <summary>Set when the message lives in another folder than the one being viewed, e.g. "Sent".</summary>
    public string? FolderLabel { get; }
    public bool HasFolderLabel => FolderLabel is not null;

    public string Sender => Message.SenderDisplay;
    public string Preview => Message.Preview;
    public string To => Message.To;
    public string Date => Message.Date.ToLocalTime().ToString("g", CultureInfo.CurrentCulture);
    public string Initials => MessageItemViewModel.InitialsFor(Sender);

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsCollapsed))]
    private bool _isExpanded;

    public bool IsCollapsed => !IsExpanded;

    [RelayCommand]
    private async Task ToggleAsync()
    {
        IsExpanded = !IsExpanded;
        if (IsExpanded && Body.Full is null)
            await Body.LoadAsync(Message);
    }

    public async Task ExpandAsync()
    {
        IsExpanded = true;
        if (Body.Full is null)
            await Body.LoadAsync(Message);
    }
}
