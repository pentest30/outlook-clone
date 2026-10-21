using Outlook.Core.Models;
using Outlook.Core.Search;

namespace Outlook.Core.Services;

/// <summary>The local cache (Outlook's .ost equivalent). Everything the UI shows comes from here.</summary>
public interface IMailStore
{
    Task<IReadOnlyList<Account>> GetAccountsAsync(CancellationToken ct = default);
    Task<Account> AddAccountAsync(Account account, CancellationToken ct = default);

    Task<IReadOnlyList<MailFolder>> GetFoldersAsync(int accountId, CancellationToken ct = default);
    Task<IReadOnlyList<MailMessage>> GetMessagesAsync(int folderId, int take = 500, CancellationToken ct = default);
    Task<MailMessage?> GetMessageAsync(int messageId, CancellationToken ct = default);

    /// <summary>Full-text search over the cache, newest first. folderId null searches every folder of the account.</summary>
    Task<IReadOnlyList<MailMessage>> SearchAsync(int accountId, int? folderId, MailSearchQuery query, int take = 300, CancellationToken ct = default);

    /// <summary>Every cached message of these conversations, in any folder (a thread spans Inbox and Sent).</summary>
    Task<IReadOnlyList<MailMessage>> GetThreadMessagesAsync(int accountId, IReadOnlyCollection<string> threadIds, CancellationToken ct = default);

    /// <summary>The account's categories, creating Outlook's six defaults the first time.</summary>
    Task<IReadOnlyList<Category>> GetCategoriesAsync(int accountId, CancellationToken ct = default);
    Task<Category> AddCategoryAsync(int accountId, string name, string color, CancellationToken ct = default);

    /// <summary>Attachment metadata only; <see cref="MailAttachment.Data"/> is left empty.</summary>
    Task<IReadOnlyList<MailAttachment>> GetAttachmentsAsync(int messageId, CancellationToken ct = default);
    Task<byte[]> GetAttachmentDataAsync(int attachmentId, CancellationToken ct = default);
}

public interface ICredentialStore
{
    void Save(int accountId, string secret);
    string? Load(int accountId);
    void Delete(int accountId);
}
