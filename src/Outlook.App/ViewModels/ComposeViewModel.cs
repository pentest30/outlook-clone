using System.Collections.ObjectModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Compose;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

/// <summary>
/// One compose window (new, reply, reply all, forward, or an opened draft). Drafts auto-save to the
/// server every 30 seconds and when the window closes, always replacing the same draft.
/// </summary>
public partial class ComposeViewModel : ObservableObject
{
    private static readonly TimeSpan AutoSaveInterval = TimeSpan.FromSeconds(30);

    private readonly IMailService _mail;
    private readonly Account _account;
    private readonly ComposeMode _mode;
    private readonly MailMessage? _original;
    private readonly string? _inReplyTo;
    private readonly string? _references;
    private readonly DispatcherTimer _autoSave;
    private bool _draftExists;
    private bool _finished;   // sent or discarded: never save again
    private string _savedSnapshot = "";

    public ComposeViewModel(IMailService mail, Account account, ComposeMode mode, MailMessage? original,
        ComposeSeed seed, IEnumerable<OutgoingAttachment> attachments, string? draftMessageId)
    {
        _mail = mail;
        _account = account;
        _mode = mode;
        _original = original;
        _inReplyTo = seed.InReplyTo;
        _references = seed.References;
        _draftExists = draftMessageId is not null;
        MessageId = draftMessageId ?? $"{Guid.NewGuid():N}@{AddressList.DomainOf(account.Email)}";

        To = seed.To;
        Cc = seed.Cc;
        ShowCc = true;
        Subject = seed.Subject;
        Body = seed.Body;
        foreach (var a in attachments)
            Attachments.Add(new AttachmentViewModel(a));
        _savedSnapshot = Snapshot();

        _autoSave = new DispatcherTimer { Interval = AutoSaveInterval };
        _autoSave.Tick += async (_, _) => await AutoSaveAsync();
        _autoSave.Start();
    }

    /// <summary>
    /// Raised after a send, save or discard so the main window can refresh Sent and Drafts. Carries a
    /// message for the main status bar when something worth knowing happened (the compose window is gone by then).
    /// </summary>
    public event Action<string?>? Completed;
    public event Action? CloseRequested;

    public string MessageId { get; }
    public string From => $"{_account.DisplayName} <{_account.Email}>";
    public ObservableCollection<AttachmentViewModel> Attachments { get; } = [];

    /// <summary>True when the body should get focus first (replies), false for To (new mail, forwards).</summary>
    public bool FocusBody => _mode is ComposeMode.Reply or ComposeMode.ReplyAll;

    [ObservableProperty] private string _to = "";
    [ObservableProperty] private string _cc = "";
    [ObservableProperty] private string _bcc = "";
    [ObservableProperty] private bool _showCc;
    [ObservableProperty] private bool _showBcc;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(Title))] private string _subject = "";
    [ObservableProperty] private string _body = "";
    [ObservableProperty] private string _status = "";

    /// <summary>
    /// Compared with the last saved state rather than tracked from edits: text boxes normalise line
    /// endings when they load, which would otherwise make every untouched reply look edited.
    /// </summary>
    public bool IsDirty => Snapshot() != _savedSnapshot;

    private string Snapshot() => string.Join('|',
        To.Trim(), Cc.Trim(), Bcc.Trim(), Subject.Trim(), Body.Replace("\r\n", "\n").TrimEnd(),
        string.Join(',', Attachments.Select(a => $"{a.FileName}:{a.Size}")));

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SendCommand), nameof(SaveDraftCommand), nameof(DiscardCommand))]
    private bool _isBusy;

    public string Title => string.IsNullOrWhiteSpace(Subject) ? "Untitled - Message" : $"{Subject} - Message";

    private bool CanAct() => !IsBusy;

    /// <summary>
    /// Set by the main window to hold the message for a few seconds ("Undo send"). The window
    /// closes immediately and the outbox calls <see cref="SendNowAsync"/> when the delay is over.
    /// </summary>
    public Action<ComposeViewModel>? QueueSend { get; set; }

    [RelayCommand(CanExecute = nameof(CanAct))]
    private async Task SendAsync()
    {
        // Catch typos while the window is still open, not seconds later from the outbox.
        if (RecipientProblem() is { } problem)
        {
            Status = problem;
            return;
        }

        if (QueueSend is { } queue)
        {
            Finish();
            queue(this);
            CloseRequested?.Invoke();
            return;
        }

        if (await SendNowAsync())
            CloseRequested?.Invoke();
    }

    /// <summary>Sends, then does the follow-ups. On failure the message stays editable and Status says why.</summary>
    public async Task<bool> SendNowAsync()
    {
        IsBusy = true;
        Status = "Sending…";
        SendReport report;
        try
        {
            report = await _mail.SendAsync(_account, Build());
        }
        catch (Exception ex)
        {
            Status = $"Couldn't send: {ex.Message}";
            IsBusy = false;
            Reopen();
            return false;
        }

        Finish();
        // Follow-ups: the message is already sent, so a failure here is not worth blocking on.
        try
        {
            if (_original is not null && _mode is ComposeMode.Reply or ComposeMode.ReplyAll or ComposeMode.Forward)
                await _mail.MarkRespondedAsync(_original, forwarded: _mode == ComposeMode.Forward);
            if (_draftExists)
                await _mail.DeleteDraftAsync(_account, MessageId);
        }
        catch (Exception)
        {
            // The next sync shows the true server state.
        }
        IsBusy = false;
        Completed?.Invoke(report.Problem);
        return true;
    }

    /// <summary>Back to editing after Undo or a failed send.</summary>
    public void Reopen()
    {
        _finished = false;
        _autoSave.Start();
    }

    private string? RecipientProblem()
    {
        var entries = new[] { To, Cc, Bcc }.SelectMany(AddressList.Split).ToList();
        if (entries.Count == 0)
            return "Add at least one recipient.";
        var bad = entries.FirstOrDefault(e => !MimeKit.MailboxAddress.TryParse(e, out var m) || !m.Address.Contains('@'));
        return bad is null ? null : $"\"{bad}\" isn't a valid email address.";
    }

    [RelayCommand(CanExecute = nameof(CanAct))]
    private async Task SaveDraftAsync()
    {
        IsBusy = true;
        try
        {
            await SaveDraftCoreAsync();
            Completed?.Invoke(null);
        }
        catch (Exception ex)
        {
            Status = $"Couldn't save the draft: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand(CanExecute = nameof(CanAct))]
    private async Task DiscardAsync()
    {
        Finish();
        if (_draftExists)
        {
            try
            {
                await _mail.DeleteDraftAsync(_account, MessageId);
            }
            catch (Exception)
            {
                // Leaves a stray draft at worst.
            }
            Completed?.Invoke(null);
        }
        CloseRequested?.Invoke();
    }

    [RelayCommand]
    private void ToggleBcc() => ShowBcc = !ShowBcc;

    [RelayCommand]
    private void RemoveAttachment(AttachmentViewModel attachment) => Attachments.Remove(attachment);

    public void AddAttachment(string fileName, byte[] data)
    {
        Attachments.Add(new AttachmentViewModel(new OutgoingAttachment(fileName, MimeKit.MimeTypes.GetMimeType(fileName), data)));
    }

    /// <summary>Whether closing the window should first save a draft (unsent edits).</summary>
    public bool NeedsSaveOnClose => !_finished && IsDirty;

    public async Task SaveOnCloseAsync()
    {
        _autoSave.Stop();
        if (!NeedsSaveOnClose) return;
        try
        {
            await SaveDraftCoreAsync();
            Completed?.Invoke(null);
        }
        catch (Exception)
        {
            // Closing anyway; an unsaved draft is lost, as with a crash.
        }
        _finished = true;
    }

    private async Task AutoSaveAsync()
    {
        if (!NeedsSaveOnClose || IsBusy) return;
        try
        {
            await SaveDraftCoreAsync();
        }
        catch (Exception ex)
        {
            Status = $"Couldn't save the draft: {ex.Message}";
        }
    }

    private async Task SaveDraftCoreAsync()
    {
        Status = "Saving draft…";
        await _mail.SaveDraftAsync(_account, Build());
        _draftExists = true;
        _savedSnapshot = Snapshot();
        Status = $"Draft saved at {DateTime.Now:t}";
    }

    private void Finish()
    {
        _finished = true;
        _autoSave.Stop();
    }

    private OutgoingMessage Build() => new(
        [To], [Cc], [Bcc], Subject, Body,
        InReplyTo: _inReplyTo,
        References: _references,
        Attachments: Attachments.Select(a => a.Outgoing!).ToList(),
        MessageId: MessageId);
}
