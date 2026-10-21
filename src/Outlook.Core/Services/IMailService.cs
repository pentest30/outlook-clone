using Outlook.Core.Models;

namespace Outlook.Core.Services;

public record SyncProgress(string FolderName, int FolderIndex, int FolderCount);

/// <summary>Outcome of a send that reached the SMTP server. Problem is set when a follow-up step failed.</summary>
public record SendReport(bool SavedToSent, string? Problem);

public class FolderChangedEventArgs(int accountId, string folderFullName) : EventArgs
{
    public int AccountId { get; } = accountId;
    public string FolderFullName { get; } = folderFullName;
}

/// <summary>
/// Talks to the mail server and keeps the local cache in step with it.
/// The UI reads from <see cref="IMailStore"/> and calls this for anything that touches the server.
/// </summary>
public interface IMailService
{
    /// <summary>Raised (on a background thread) when push notifications report new or changed mail.</summary>
    event EventHandler<FolderChangedEventArgs>? FolderChanged;

    /// <summary>Connects and logs in, throwing if the settings or password are wrong.</summary>
    Task VerifyAsync(Account account, string password, CancellationToken ct = default);

    Task SyncAccountAsync(Account account, IProgress<SyncProgress>? progress = null, CancellationToken ct = default);

    Task SyncFolderAsync(Account account, MailFolder folder, CancellationToken ct = default);

    /// <summary>
    /// Holds an IMAP IDLE session on the Inbox and raises <see cref="FolderChanged"/> until cancelled.
    /// Returns immediately when the server has no IDLE support; periodic sync covers that case.
    /// </summary>
    Task RunPushAsync(Account account, CancellationToken ct);

    /// <summary>Downloads the full message, caching body and attachments.</summary>
    Task<MailMessage> DownloadBodyAsync(MailMessage message, CancellationToken ct = default);

    Task SetReadAsync(MailMessage message, bool isRead, CancellationToken ct = default);

    Task SetFlaggedAsync(MailMessage message, bool isFlagged, CancellationToken ct = default);

    /// <summary>Marks the original as answered (reply) or forwarded once the response is sent.</summary>
    Task MarkRespondedAsync(MailMessage original, bool forwarded, CancellationToken ct = default);

    /// <summary>Adds or removes a category keyword on the server and in the cache.</summary>
    Task SetCategoryAsync(IReadOnlyList<MailMessage> messages, string keyword, bool add, CancellationToken ct = default);

    Task MoveAsync(IReadOnlyList<MailMessage> messages, MailFolder target, CancellationToken ct = default);

    /// <summary>Moves to Deleted Items, or deletes permanently when already there.</summary>
    Task DeleteAsync(IReadOnlyList<MailMessage> messages, CancellationToken ct = default);

    /// <summary>Moves to the Archive folder, creating it on the server if it does not exist.</summary>
    Task ArchiveAsync(IReadOnlyList<MailMessage> messages, CancellationToken ct = default);

    /// <summary>Throws only when the message was not sent; see <see cref="SendReport"/> for what happened after.</summary>
    Task<SendReport> SendAsync(Account account, OutgoingMessage message, CancellationToken ct = default);

    /// <summary>Saves to Drafts, replacing the earlier copy with the same MessageId.</summary>
    Task SaveDraftAsync(Account account, OutgoingMessage draft, CancellationToken ct = default);

    Task DeleteDraftAsync(Account account, string messageId, CancellationToken ct = default);
}
