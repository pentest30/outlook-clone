using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Core.Text;

namespace Outlook.App.ViewModels;

/// <summary>
/// The rendered body of one message: HTML, attachments, inline pictures and the remote-picture
/// block. Shared by the single-message reading pane and each card of a conversation.
/// </summary>
public partial class MessageBodyViewModel(IMailService mail, IMailStore store,
    Func<string, Task<InvitationViewModel?>>? invitationFor = null) : ObservableObject
{
    public ObservableCollection<AttachmentViewModel> Attachments { get; } = [];

    /// <summary>The full message (with body) once loaded.</summary>
    public MailMessage? Full { get; private set; }

    [ObservableProperty] private string _html = "";

    /// <summary>The message has no HTML part: drawn in the theme's colours rather than on white.</summary>
    [ObservableProperty] private bool _isPlainText;
    [ObservableProperty] private bool _isLoading;
    [ObservableProperty] private bool _hasBlockedContent;
    [ObservableProperty] private bool _allowRemoteContent;
    [ObservableProperty] private bool _hasAttachments;
    [ObservableProperty] private Func<string, byte[]?>? _inlineImageResolver;

    /// <summary>The meeting card, when the message carries an invitation, update, cancellation or reply.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasInvitation))]
    private InvitationViewModel? _invitation;

    public bool HasInvitation => Invitation is not null;

    /// <summary>Downloads the body if needed and renders it. Returns false if the message is gone.</summary>
    public async Task<bool> LoadAsync(MailMessage message)
    {
        IsLoading = true;
        try
        {
            var full = message.IsBodyDownloaded ? await store.GetMessageAsync(message.Id) : await mail.DownloadBodyAsync(message);
            if (full is null) return false;   // moved away by a sync meanwhile
            message.IsBodyDownloaded = true;

            var attachments = await store.GetAttachmentsAsync(full.Id);
            var inline = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach (var a in attachments.Where(a => a.IsInline && a.ContentId is not null))
                inline[a.ContentId!] = await store.GetAttachmentDataAsync(a.Id);

            Attachments.Clear();
            foreach (var a in attachments.Where(a => !a.IsInline))
                Attachments.Add(new AttachmentViewModel(a, () => store.GetAttachmentDataAsync(a.Id)));
            HasAttachments = Attachments.Count > 0;
            InlineImageResolver = inline.Count == 0 ? null : id => inline.GetValueOrDefault(id);

            var prepared = EmailHtml.Prepare(full.BodyHtml, full.BodyText);
            IsPlainText = string.IsNullOrWhiteSpace(full.BodyHtml);
            AllowRemoteContent = false;
            HasBlockedContent = prepared.HasRemoteContent;
            Html = prepared.Html;
            Full = full;

            Invitation = full.CalendarIcs is { } ics && invitationFor is not null ? await invitationFor(ics) : null;
            if (Invitation is not null)
                await Invitation.AutoApplyAsync();
            return true;
        }
        catch (Exception ex)
        {
            Html = EmailHtml.FromPlainText($"Couldn't download this message: {ex.Message}");
            IsPlainText = true;
            return false;
        }
        finally
        {
            IsLoading = false;
        }
    }

    public void Clear()
    {
        Full = null;
        Html = "";
        IsPlainText = false;
        Attachments.Clear();
        HasAttachments = false;
        HasBlockedContent = false;
        AllowRemoteContent = false;
        InlineImageResolver = null;
        Invitation = null;
    }

    [RelayCommand]
    private void DownloadPictures()
    {
        AllowRemoteContent = true;
        HasBlockedContent = false;
    }
}
