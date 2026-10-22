using System.Collections.Concurrent;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Search;
using Microsoft.EntityFrameworkCore;
using MimeKit;
using MimeKit.Utils;
using Outlook.Core.Compose;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Core.Text;
using Outlook.Core.Threading;
using Outlook.Data;
using IMailService = Outlook.Core.Services.IMailService;
using SpecialFolder = Outlook.Core.Models.SpecialFolder;
using MailFolder = Outlook.Core.Models.MailFolder;

namespace Outlook.Mail;

/// <summary>
/// IMAP/SMTP implementation of <see cref="IMailService"/>. See docs/PLAN.md for the sync model.
/// </summary>
public sealed class ImapMailService(Func<MailDbContext> createContext, ICredentialStore credentials)
    : IMailService, IAsyncDisposable
{
    /// <summary>First sync of a big folder only pulls the newest messages so the app is usable quickly.</summary>
    public int InitialSyncLimit { get; init; } = 2000;

    private const int FetchBatchSize = 200;
    private const string ForwardedKeyword = "$Forwarded";

    private const MessageSummaryItems HeaderItems =
        MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope | MessageSummaryItems.Flags |
        MessageSummaryItems.InternalDate | MessageSummaryItems.Size | MessageSummaryItems.BodyStructure |
        MessageSummaryItems.PreviewText | MessageSummaryItems.References;

    /// <summary>
    /// Two sessions per account, like Outlook: a background one for sync and an interactive one, so
    /// opening a message is not stuck behind a long folder sync. Push (IDLE) holds a third, see RunPushAsync.
    /// </summary>
    private enum Lane { Sync, Interactive }

    private readonly ConcurrentDictionary<(int AccountId, Lane Lane), ImapConnection> _connections = new();

    public event EventHandler<FolderChangedEventArgs>? FolderChanged;

    public async Task ForgetAccountAsync(int accountId)
    {
        foreach (var key in _connections.Keys.Where(k => k.AccountId == accountId).ToList())
            if (_connections.TryRemove(key, out var connection))
                await connection.DisposeAsync();
    }

    public async Task VerifySmtpAsync(Account account, string password, CancellationToken ct = default)
    {
        using var smtp = new SmtpClient();
        await smtp.ConnectAsync(account.SmtpHost, account.SmtpPort, ImapConnection.ToSocketOptions(account.SmtpSecurity), ct);
        if (smtp.Capabilities.HasFlag(SmtpCapabilities.Authentication))
            await smtp.AuthenticateAsync(account.UserName, password, ct);
        await smtp.DisconnectAsync(true, ct);
    }

    public async Task VerifyAsync(Account account, string password, CancellationToken ct = default)
    {
        using var client = new ImapClient();
        await client.ConnectAsync(account.ImapHost, account.ImapPort, ImapConnection.ToSocketOptions(account.ImapSecurity), ct);
        await client.AuthenticateAsync(account.UserName, password, ct);
        await client.DisconnectAsync(true, ct);
    }

    // ───────────────────────────── Sync ─────────────────────────────

    public async Task SyncAccountAsync(Account account, IProgress<SyncProgress>? progress = null, CancellationToken ct = default)
    {
        await ConnectionFor(account, Lane.Sync).UseAsync(async client =>
        {
            var remoteFolders = await SyncFolderListAsync(client, account, ct);
            for (var i = 0; i < remoteFolders.Count; i++)
            {
                var (remote, localId) = remoteFolders[i];
                progress?.Report(new SyncProgress(remote.Name, i, remoteFolders.Count));
                await SyncFolderAsync(remote, localId, ct);
            }
        }, ct);
    }

    public Task SyncFolderAsync(Account account, MailFolder folder, CancellationToken ct = default) =>
        ConnectionFor(account, Lane.Sync).UseAsync(async client =>
        {
            var remote = await client.GetFolderAsync(folder.FullName, ct);
            await SyncFolderAsync(remote, folder.Id, ct);
        }, ct);

    /// <summary>Mirrors the server's folder tree into the cache and returns the selectable folders.</summary>
    private async Task<List<(IMailFolder Remote, int LocalId)>> SyncFolderListAsync(ImapClient client, Account account, CancellationToken ct)
    {
        var remote = (await client.GetFoldersAsync(client.PersonalNamespaces[0], cancellationToken: ct))
            .Where(f => !f.Attributes.HasFlag(FolderAttributes.NonExistent))
            .ToList();
        if (!remote.Any(f => f.FullName == client.Inbox.FullName))
            remote.Insert(0, client.Inbox);

        await using var db = createContext();
        var local = await db.Folders.Where(f => f.AccountId == account.Id).ToDictionaryAsync(f => f.FullName, ct);

        foreach (var r in remote)
        {
            if (!local.TryGetValue(r.FullName, out var folder))
            {
                folder = new MailFolder { AccountId = account.Id, FullName = r.FullName };
                db.Folders.Add(folder);
                local[r.FullName] = folder;
            }
            folder.DisplayName = r.FullName == client.Inbox.FullName ? "Inbox" : r.Name;
            folder.ParentFullName = string.IsNullOrEmpty(r.ParentFolder?.FullName) ? null : r.ParentFolder.FullName;
            folder.Special = ClassifyFolder(r, client.Inbox);
        }

        var remoteNames = remote.Select(r => r.FullName).ToHashSet();
        db.Folders.RemoveRange(local.Values.Where(f => !remoteNames.Contains(f.FullName)));
        await db.SaveChangesAsync(ct);

        return remote
            .Where(r => !r.Attributes.HasFlag(FolderAttributes.NoSelect))
            .Select(r => (r, local[r.FullName].Id))
            .ToList();
    }

    internal static SpecialFolder ClassifyFolder(IMailFolder folder, IMailFolder inbox)
    {
        if (folder.FullName == inbox.FullName) return SpecialFolder.Inbox;
        var a = folder.Attributes;
        if (a.HasFlag(FolderAttributes.Drafts)) return SpecialFolder.Drafts;
        if (a.HasFlag(FolderAttributes.Sent)) return SpecialFolder.Sent;
        if (a.HasFlag(FolderAttributes.Trash)) return SpecialFolder.Trash;
        if (a.HasFlag(FolderAttributes.Junk)) return SpecialFolder.Junk;
        if (a.HasFlag(FolderAttributes.Archive)) return SpecialFolder.Archive;
        return ClassifyByName(folder.Name);
    }

    /// <summary>Fallback for servers without SPECIAL-USE, matching the names common clients create.</summary>
    internal static SpecialFolder ClassifyByName(string name) => name.ToLowerInvariant() switch
    {
        "drafts" or "draft" => SpecialFolder.Drafts,
        "sent" or "sent items" or "sent mail" or "sent messages" => SpecialFolder.Sent,
        "trash" or "deleted items" or "deleted messages" or "bin" => SpecialFolder.Trash,
        "junk" or "junk e-mail" or "junk email" or "spam" => SpecialFolder.Junk,
        "archive" or "archives" => SpecialFolder.Archive,
        _ => SpecialFolder.None,
    };

    private async Task SyncFolderAsync(IMailFolder remote, int localFolderId, CancellationToken ct)
    {
        await remote.OpenAsync(FolderAccess.ReadOnly, ct);
        await using var db = createContext();
        var folder = await db.Folders.SingleAsync(f => f.Id == localFolderId, ct);

        if (folder.UidValidity != remote.UidValidity)
        {
            // The server renumbered the folder, so cached UIDs point at the wrong messages.
            await db.Messages.Where(m => m.FolderId == folder.Id).ExecuteDeleteAsync(ct);
            folder.UidValidity = remote.UidValidity;
        }

        var serverUids = remote.Count == 0 ? [] : await remote.SearchAsync(SearchQuery.All, ct);
        var serverSet = serverUids.Select(u => u.Id).ToHashSet();
        var cached = await db.Messages.Where(m => m.FolderId == folder.Id).ToDictionaryAsync(m => m.Uid, ct);

        // 1. Messages deleted or moved elsewhere on the server.
        db.Messages.RemoveRange(cached.Values.Where(m => !serverSet.Contains(m.Uid)));

        // 2. Flag changes (read, flagged, answered) on messages we already have.
        var known = serverUids.Where(u => cached.ContainsKey(u.Id)).ToList();
        foreach (var batch in known.Chunk(FetchBatchSize * 5))
        {
            foreach (var s in await remote.FetchAsync(batch, new FetchRequest(MessageSummaryItems.UniqueId | MessageSummaryItems.Flags), ct))
                ApplyFlags(cached[s.UniqueId.Id], s);
        }

        // 3. New messages, newest first, capped on the first sync.
        var fresh = serverUids.Where(u => !cached.ContainsKey(u.Id)).OrderByDescending(u => u.Id).ToList();
        if (cached.Count == 0 && fresh.Count > InitialSyncLimit)
            fresh = fresh.Take(InitialSyncLimit).ToList();
        // New mail feeds the auto-complete list: who wrote to us, and whom we wrote to (Sent).
        var account = fresh.Count > 0 ? await db.Accounts.SingleAsync(a => a.Id == folder.AccountId, ct) : null;
        var book = account is null ? null : new AddressBook(db, account);
        var isSent = folder.Special == SpecialFolder.Sent;
        foreach (var batch in fresh.Chunk(FetchBatchSize))
        {
            foreach (var s in await remote.FetchAsync(batch, new FetchRequest(HeaderItems), ct))
            {
                var message = ToMessage(folder.Id, s);
                db.Messages.Add(message);
                book?.Record(isSent ? AddressBook.People(message.To, message.Cc) : [(message.FromName, message.FromAddress)],
                    sent: isSent, message.Date);
            }
        }

        folder.LastSyncedAt = DateTimeOffset.UtcNow;
        await db.SaveChangesAsync(ct);
        await RecountAsync(db, [folder.Id], ct);
    }

    internal static MailMessage ToMessage(int folderId, IMessageSummary s)
    {
        var env = s.Envelope;
        var from = env?.From.Mailboxes.FirstOrDefault();
        var references = s.References is { Count: > 0 } refs ? string.Join(' ', refs) : null;
        var subject = env?.Subject ?? "";

        var message = new MailMessage
        {
            FolderId = folderId,
            Uid = s.UniqueId.Id,
            MessageIdHeader = env?.MessageId,
            InReplyTo = env?.InReplyTo,
            References = references,
            ThreadId = ThreadResolver.Resolve(env?.MessageId, env?.InReplyTo, references, subject),
            Subject = subject,
            FromName = from?.Name ?? "",
            FromAddress = from?.Address ?? "",
            To = FormatAddresses(env?.To),
            Cc = FormatAddresses(env?.Cc),
            ReplyTo = FormatAddresses(env?.ReplyTo),
            Date = env?.Date ?? s.InternalDate ?? DateTimeOffset.UtcNow,
            Size = s.Size ?? 0,
            Preview = (s.PreviewText ?? "").Trim(),
            HasAttachments = s.Attachments?.Any() ?? false,
            HasCalendar = s.BodyParts?.Any(IsCalendarPart) ?? false,
        };
        ApplyFlags(message, s);
        return message;
    }

    private static bool IsCalendarPart(BodyPartBasic part) =>
        part.ContentType.IsMimeType("text", "calendar") || part.ContentType.IsMimeType("application", "ics")
        || (part.FileName?.EndsWith(".ics", StringComparison.OrdinalIgnoreCase) ?? false);

    /// <summary>The invitation in a message: the text/calendar alternative, or an attached .ics.</summary>
    internal static string? FindCalendar(MimeMessage mime)
    {
        foreach (var part in mime.BodyParts.OfType<MimePart>())
        {
            var isCalendar = part.ContentType.IsMimeType("text", "calendar") || part.ContentType.IsMimeType("application", "ics")
                             || (part.FileName?.EndsWith(".ics", StringComparison.OrdinalIgnoreCase) ?? false);
            if (!isCalendar) continue;
            if (part is TextPart text) return text.Text;
            using var stream = new MemoryStream();
            part.Content?.DecodeTo(stream);
            return System.Text.Encoding.UTF8.GetString(stream.ToArray());
        }
        return null;
    }

    private static void ApplyFlags(MailMessage m, IMessageSummary s)
    {
        var flags = s.Flags ?? MessageFlags.None;
        m.IsRead = flags.HasFlag(MessageFlags.Seen);
        m.IsFlagged = flags.HasFlag(MessageFlags.Flagged);
        m.IsAnswered = flags.HasFlag(MessageFlags.Answered);
        m.IsDraft = flags.HasFlag(MessageFlags.Draft);
        m.IsForwarded = s.Keywords?.Contains(ForwardedKeyword) ?? false;
        m.CategoryKeywords = string.Join(' ', (s.Keywords ?? Enumerable.Empty<string>())
            .Where(k => k.StartsWith(Category.KeywordPrefix, StringComparison.Ordinal))
            .Order(StringComparer.Ordinal));
    }

    /// <summary>
    /// "Name &lt;address&gt;" for display, quoting the name only when it contains characters that would
    /// break parsing it back (e.g. "Smith, Bob"), which replies rely on.
    /// </summary>
    internal static string FormatAddresses(InternetAddressList? list) =>
        list is null ? "" : string.Join(", ", list.Mailboxes.Select(FormatMailbox));

    private static string FormatMailbox(MailboxAddress m)
    {
        if (string.IsNullOrWhiteSpace(m.Name))
            return m.Address;
        return m.Name.IndexOfAny([',', ';', '<', '>', '"', '@', ':', '(', ')', '[', ']', '\\']) >= 0
            ? m.ToString()
            : $"{m.Name} <{m.Address}>";
    }

    private static async Task RecountAsync(MailDbContext db, IEnumerable<int> folderIds, CancellationToken ct)
    {
        foreach (var id in folderIds.Distinct())
        {
            var folder = await db.Folders.SingleOrDefaultAsync(f => f.Id == id, ct);
            if (folder is null) continue;
            folder.TotalCount = await db.Messages.CountAsync(m => m.FolderId == id, ct);
            folder.UnreadCount = await db.Messages.CountAsync(m => m.FolderId == id && !m.IsRead, ct);
        }
        await db.SaveChangesAsync(ct);
    }

    // ───────────────────────────── Push (IMAP IDLE) ─────────────────────────────

    public async Task RunPushAsync(Account account, CancellationToken ct)
    {
        var retryDelay = TimeSpan.FromSeconds(5);
        while (!ct.IsCancellationRequested)
        {
            try
            {
                using var client = new ImapClient();
                await client.ConnectAsync(account.ImapHost, account.ImapPort, ImapConnection.ToSocketOptions(account.ImapSecurity), ct);
                await client.AuthenticateAsync(account.UserName, PasswordFor(account), ct);
                if (!client.Capabilities.HasFlag(ImapCapabilities.Idle))
                    return;

                var inbox = client.Inbox;
                await inbox.OpenAsync(FolderAccess.ReadOnly, ct);
                retryDelay = TimeSpan.FromSeconds(5);

                // Server notifications arrive as folder events while IDLE runs. Cancelling `done`
                // ends the IDLE so we can report the change, then we go straight back to idling.
                CancellationTokenSource? done = null;
                var changed = false;
                void OnChange(object? sender, EventArgs e)
                {
                    changed = true;
                    done?.Cancel();
                }
                inbox.CountChanged += OnChange;
                inbox.MessageExpunged += OnChange;
                inbox.MessageFlagsChanged += OnChange;

                while (!ct.IsCancellationRequested)
                {
                    // Servers may drop an IDLE after 30 minutes (RFC 2177), so re-issue it every 9.
                    using (done = new CancellationTokenSource(TimeSpan.FromMinutes(9)))
                        await client.IdleAsync(done.Token, ct);
                    done = null;

                    if (changed)
                    {
                        changed = false;
                        FolderChanged?.Invoke(this, new FolderChangedEventArgs(account.Id, inbox.FullName));
                    }
                }
                await client.DisconnectAsync(true, CancellationToken.None);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                return;
            }
            catch (Exception)
            {
                // Network blip or server restart: back off and reconnect. Periodic sync still runs meanwhile.
                try { await Task.Delay(retryDelay, ct); }
                catch (OperationCanceledException) { return; }
                retryDelay = TimeSpan.FromTicks(Math.Min(retryDelay.Ticks * 2, TimeSpan.FromMinutes(5).Ticks));
            }
        }
    }

    // ───────────────────────────── Reading ─────────────────────────────

    public async Task<MailMessage> DownloadBodyAsync(MailMessage message, CancellationToken ct = default)
    {
        var (account, folderName) = await LocateAsync(message.FolderId, ct);
        var mime = await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
        {
            var folder = await client.GetFolderAsync(folderName, ct);
            await folder.OpenAsync(FolderAccess.ReadOnly, ct);
            // MailKit fetches with BODY.PEEK, so opening a message does not mark it read on the server.
            return await folder.GetMessageAsync(new UniqueId(message.Uid), ct);
        }, ct);

        await using var db = createContext();
        var cached = await db.Messages.SingleAsync(m => m.Id == message.Id, ct);
        cached.BodyText = mime.TextBody;
        cached.BodyHtml = mime.HtmlBody;
        cached.SearchText = mime.TextBody ?? (mime.HtmlBody is { } html ? HtmlText.ToPlainText(html) : null);
        cached.CalendarIcs = FindCalendar(mime);
        cached.HasCalendar = cached.CalendarIcs is not null;

        await db.Attachments.Where(a => a.MessageId == cached.Id).ExecuteDeleteAsync(ct);
        var attachments = ExtractAttachments(mime);
        foreach (var a in attachments)
            a.MessageId = cached.Id;
        db.Attachments.AddRange(attachments);
        cached.HasAttachments = attachments.Any(a => !a.IsInline);
        cached.IsBodyDownloaded = true;

        await db.SaveChangesAsync(ct);
        cached.Attachments = [];   // callers load attachments through IMailStore, without the bytes
        return cached;
    }

    /// <summary>Every part that is not the text/HTML body: files, attached e-mails and inline pictures.</summary>
    internal static List<MailAttachment> ExtractAttachments(MimeMessage mime)
    {
        var result = new List<MailAttachment>();
        foreach (var entity in mime.BodyParts)
        {
            using var stream = new MemoryStream();
            string fileName;
            switch (entity)
            {
                case MessagePart { Message: null }:
                    continue;
                case MessagePart attachedMail:
                    attachedMail.Message.WriteTo(stream);
                    fileName = entity.ContentDisposition?.FileName ?? $"{attachedMail.Message.Subject ?? "message"}.eml";
                    break;
                case MimePart part when part is TextPart && !part.IsAttachment:
                    continue;   // the body itself
                case MimePart part:
                    part.Content?.DecodeTo(stream);
                    fileName = part.FileName ?? part.ContentId ?? "attachment";
                    break;
                default:
                    continue;
            }

            var isInline = !entity.IsAttachment && entity.ContentId is not null;
            result.Add(new MailAttachment
            {
                FileName = SafeFileName(fileName),
                ContentType = entity.ContentType.MimeType,
                ContentId = entity.ContentId,
                IsInline = isInline,
                Size = stream.Length,
                Data = stream.ToArray(),
            });
        }
        return result;
    }

    /// <summary>Attachment names come from the sender, so never let one carry a path.</summary>
    private static string SafeFileName(string name)
    {
        var file = Path.GetFileName(name.Replace('\\', '/'));
        foreach (var c in Path.GetInvalidFileNameChars())
            file = file.Replace(c, '_');
        return string.IsNullOrWhiteSpace(file) ? "attachment" : file;
    }

    // ───────────────────────────── Flags ─────────────────────────────

    public Task SetReadAsync(MailMessage message, bool isRead, CancellationToken ct = default) =>
        StoreFlagAsync(message, new StoreFlagsRequest(isRead ? StoreAction.Add : StoreAction.Remove, MessageFlags.Seen),
            m => m.IsRead = isRead, m => m.IsRead = !isRead, ct);

    public Task SetFlaggedAsync(MailMessage message, bool isFlagged, CancellationToken ct = default) =>
        StoreFlagAsync(message, new StoreFlagsRequest(isFlagged ? StoreAction.Add : StoreAction.Remove, MessageFlags.Flagged),
            m => m.IsFlagged = isFlagged, m => m.IsFlagged = !isFlagged, ct);

    public Task MarkRespondedAsync(MailMessage original, bool forwarded, CancellationToken ct = default) =>
        forwarded
            ? StoreFlagAsync(original, new StoreFlagsRequest(StoreAction.Add, new HashSet<string> { ForwardedKeyword }),
                m => m.IsForwarded = true, m => m.IsForwarded = false, ct)
            : StoreFlagAsync(original, new StoreFlagsRequest(StoreAction.Add, MessageFlags.Answered),
                m => m.IsAnswered = true, m => m.IsAnswered = false, ct);

    /// <summary>Optimistic update: write the cache first, then the server, and undo the cache if the server refuses.</summary>
    private async Task StoreFlagAsync(MailMessage message, StoreFlagsRequest request,
        Action<MailMessage> apply, Action<MailMessage> revert, CancellationToken ct)
    {
        await UpdateCachedAsync(message.Id, apply, ct);
        try
        {
            var (account, folderName) = await LocateAsync(message.FolderId, ct);
            request.Silent = true;
            await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
            {
                var folder = await client.GetFolderAsync(folderName, ct);
                await folder.OpenAsync(FolderAccess.ReadWrite, ct);
                await folder.StoreAsync(new UniqueId(message.Uid), request, ct);
            }, ct);
        }
        catch
        {
            await UpdateCachedAsync(message.Id, revert, CancellationToken.None);
            throw;
        }
        apply(message);
    }

    private async Task UpdateCachedAsync(int messageId, Action<MailMessage> change, CancellationToken ct)
    {
        await using var db = createContext();
        var m = await db.Messages.SingleAsync(x => x.Id == messageId, ct);
        var wasRead = m.IsRead;
        change(m);
        if (wasRead != m.IsRead)
        {
            var folder = await db.Folders.SingleAsync(f => f.Id == m.FolderId, ct);
            folder.UnreadCount += m.IsRead ? -1 : 1;
        }
        await db.SaveChangesAsync(ct);
    }

    // ───────────────────────────── Categories ─────────────────────────────

    public async Task SetCategoryAsync(IReadOnlyList<MailMessage> messages, string keyword, bool add, CancellationToken ct = default)
    {
        foreach (var group in messages.GroupBy(m => m.FolderId))
        {
            var (account, folderName) = await LocateAsync(group.Key, ct);
            var uids = group.Select(m => new UniqueId(m.Uid)).ToList();
            await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
            {
                var folder = await client.GetFolderAsync(folderName, ct);
                await folder.OpenAsync(FolderAccess.ReadWrite, ct);
                // Servers list "\*" in PERMANENTFLAGS when they accept new keywords; without it the
                // category would vanish on the next sync, so say so instead of pretending it worked.
                if (!folder.PermanentFlags.HasFlag(MessageFlags.UserDefined) && !folder.PermanentKeywords.Contains(keyword))
                    throw new NotSupportedException("This mail server doesn't store categories (custom keywords).");
                var request = new StoreFlagsRequest(add ? StoreAction.Add : StoreAction.Remove, new HashSet<string> { keyword }) { Silent = true };
                await folder.StoreAsync(uids, request, ct);
            }, ct);

            await using var db = createContext();
            var ids = group.Select(m => m.Id).ToList();
            foreach (var row in await db.Messages.Where(m => ids.Contains(m.Id)).ToListAsync(ct))
            {
                var set = row.CategoryKeywordList.ToHashSet();
                if (add) set.Add(keyword); else set.Remove(keyword);
                row.CategoryKeywords = string.Join(' ', set.Order(StringComparer.Ordinal));
            }
            await db.SaveChangesAsync(ct);
        }
    }

    // ───────────────────────────── Move / delete / archive ─────────────────────────────

    public async Task MoveAsync(IReadOnlyList<MailMessage> messages, MailFolder target, CancellationToken ct = default)
    {
        foreach (var group in messages.GroupBy(m => m.FolderId).Where(g => g.Key != target.Id))
        {
            var (account, sourceName) = await LocateAsync(group.Key, ct);
            var uids = group.Select(m => new UniqueId(m.Uid)).ToList();
            var map = await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
            {
                var source = await client.GetFolderAsync(sourceName, ct);
                await source.OpenAsync(FolderAccess.ReadWrite, ct);
                var destination = await client.GetFolderAsync(target.FullName, ct);
                // MailKit uses MOVE when the server has it, otherwise COPY + delete + expunge.
                return await source.MoveToAsync(uids, destination, ct);
            }, ct);
            await ApplyMoveToCacheAsync(group.ToList(), target.Id, map, ct);
        }
    }

    /// <summary>
    /// With UIDPLUS the server tells us each message's new UID, so the cached row (body and
    /// attachments included) simply changes folder. Otherwise drop it; the next sync of the target fetches it.
    /// </summary>
    private async Task ApplyMoveToCacheAsync(List<MailMessage> moved, int targetId, UniqueIdMap map, CancellationToken ct)
    {
        await using var db = createContext();
        var target = await db.Folders.SingleAsync(f => f.Id == targetId, ct);
        var newUids = map.Source.Zip(map.Destination).ToDictionary(p => p.First.Id, p => p.Second);
        var ids = moved.Select(m => m.Id).ToList();
        var rows = await db.Messages.Where(m => ids.Contains(m.Id)).ToListAsync(ct);
        var sourceIds = rows.Select(r => r.FolderId).Distinct().ToList();

        foreach (var row in rows)
        {
            var canRelocate = newUids.TryGetValue(row.Uid, out var newUid)
                              && target.UidValidity != 0 && newUid.Validity == target.UidValidity
                              && !await db.Messages.AnyAsync(m => m.FolderId == targetId && m.Uid == newUid.Id, ct);
            if (canRelocate)
            {
                row.FolderId = targetId;
                row.Uid = newUid.Id;
            }
            else
            {
                db.Messages.Remove(row);
            }
        }
        await db.SaveChangesAsync(ct);
        await RecountAsync(db, [.. sourceIds, targetId], ct);
    }

    /// <summary>
    /// Moves to Trash (created if the account has none, so Delete is never silently permanent), or
    /// deletes for good when the messages are already in Trash.
    /// </summary>
    public async Task DeleteAsync(IReadOnlyList<MailMessage> messages, CancellationToken ct = default)
    {
        foreach (var group in messages.GroupBy(m => m.FolderId))
        {
            MailFolder source;
            await using (var db = createContext())
                source = await db.Folders.AsNoTracking().Include(f => f.Account).SingleAsync(f => f.Id == group.Key, ct);

            if (source.Special == SpecialFolder.Trash)
            {
                await DeletePermanentlyAsync(source.Account!, source, group.ToList(), ct);
                continue;
            }
            var trash = await EnsureFolderAsync(source.Account!, SpecialFolder.Trash, "Trash", ct);
            await MoveAsync(group.ToList(), trash, ct);
        }
    }

    private async Task DeletePermanentlyAsync(Account account, MailFolder folder, List<MailMessage> messages, CancellationToken ct)
    {
        var uids = messages.Select(m => new UniqueId(m.Uid)).ToList();
        await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
        {
            var remote = await client.GetFolderAsync(folder.FullName, ct);
            await remote.OpenAsync(FolderAccess.ReadWrite, ct);
            await ExpungeAsync(client, remote, uids, ct);
        }, ct);

        await using var db = createContext();
        var ids = messages.Select(m => m.Id).ToList();
        await db.Messages.Where(m => ids.Contains(m.Id)).ExecuteDeleteAsync(ct);
        await RecountAsync(db, [folder.Id], ct);
    }

    /// <summary>Flags \Deleted and expunges just those UIDs (UIDPLUS), or the whole folder's deleted set otherwise.</summary>
    private static async Task ExpungeAsync(ImapClient client, IMailFolder folder, IList<UniqueId> uids, CancellationToken ct)
    {
        if (uids.Count == 0) return;
        await folder.StoreAsync(uids, new StoreFlagsRequest(StoreAction.Add, MessageFlags.Deleted) { Silent = true }, ct);
        if (client.Capabilities.HasFlag(ImapCapabilities.UidPlus))
            await folder.ExpungeAsync(uids, ct);
        else
            await folder.ExpungeAsync(ct);
    }

    public async Task ArchiveAsync(IReadOnlyList<MailMessage> messages, CancellationToken ct = default)
    {
        foreach (var group in messages.GroupBy(m => m.FolderId))
        {
            var (account, _) = await LocateAsync(group.Key, ct);
            var archive = await EnsureFolderAsync(account, SpecialFolder.Archive, "Archive", ct);
            await MoveAsync(group.ToList(), archive, ct);
        }
    }

    /// <summary>Returns the account's special folder, creating it on the server (as Outlook does) if missing.</summary>
    private async Task<MailFolder> EnsureFolderAsync(Account account, SpecialFolder special, string name, CancellationToken ct)
    {
        await using var db = createContext();
        var existing = await db.Folders.FirstOrDefaultAsync(f => f.AccountId == account.Id && f.Special == special, ct);
        if (existing is not null)
            return existing;

        var fullName = await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
        {
            var root = client.GetFolder(client.PersonalNamespaces[0]);
            var siblings = await root.GetSubfoldersAsync(false, ct);
            var folder = siblings.FirstOrDefault(f => f.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                         ?? await root.CreateAsync(name, true, ct)
                         ?? throw new InvalidOperationException($"The server did not create the {name} folder.");
            return folder.FullName;
        }, ct);

        var created = await db.Folders.FirstOrDefaultAsync(f => f.AccountId == account.Id && f.FullName == fullName, ct);
        if (created is null)
        {
            created = new MailFolder { AccountId = account.Id, FullName = fullName, DisplayName = name };
            db.Folders.Add(created);
        }
        created.Special = special;
        await db.SaveChangesAsync(ct);
        return created;
    }

    // ───────────────────────────── Sending and drafts ─────────────────────────────

    /// <summary>
    /// Throws only if the message did not go out. Once SMTP accepts it, later problems (saving the
    /// copy to Sent) are reported in the result instead: an exception there would make the user
    /// press Send again and deliver a duplicate.
    /// </summary>
    public async Task<SendReport> SendAsync(Account account, OutgoingMessage outgoing, CancellationToken ct = default)
    {
        var mime = BuildMime(account, outgoing);

        using (var smtp = new SmtpClient())
        {
            await smtp.ConnectAsync(account.SmtpHost, account.SmtpPort, ImapConnection.ToSocketOptions(account.SmtpSecurity), ct);
            if (!string.IsNullOrEmpty(account.UserName) && smtp.Capabilities.HasFlag(SmtpCapabilities.Authentication))
                await smtp.AuthenticateAsync(account.UserName, PasswordFor(account), ct);
            await smtp.SendAsync(mime, ct);
            await RememberRecipientsAsync(account, mime, ct);
            try { await smtp.DisconnectAsync(true, ct); }
            catch (Exception) { /* already delivered; a failed QUIT changes nothing */ }
        }

        // Keep a copy in Sent, creating the folder if the account has never had one (common for
        // mailboxes used over POP3), as Outlook and Thunderbird do. Gmail already saves a copy
        // server-side; de-duplicating that is a Phase 3 item.
        try
        {
            var sent = await EnsureFolderAsync(account, SpecialFolder.Sent, "Sent", ct);
            await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
            {
                var folder = await client.GetFolderAsync(sent.FullName, ct);
                await folder.AppendAsync(new AppendRequest(mime, MessageFlags.Seen), ct);
                await SyncFolderAsync(folder, sent.Id, ct);
            }, ct);
            return new SendReport(SavedToSent: true, Problem: null);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return new SendReport(SavedToSent: false, Problem: $"Message sent, but a copy couldn't be saved to Sent: {ex.Message}");
        }
    }

    /// <summary>Everyone just written to goes to the top of the auto-complete list.</summary>
    private async Task RememberRecipientsAsync(Account account, MimeMessage mime, CancellationToken ct)
    {
        try
        {
            await using var db = createContext();
            var book = new AddressBook(db, account);
            book.Record(mime.To.Mailboxes.Concat(mime.Cc.Mailboxes).Concat(mime.Bcc.Mailboxes).Select(m => (m.Name ?? "", m.Address)),
                sent: true, DateTimeOffset.UtcNow);
            await book.SaveAsync(ct);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            // Suggestions are a convenience; the message is already on its way.
        }
    }

    public async Task SaveDraftAsync(Account account, OutgoingMessage draft, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(draft.MessageId))
            throw new ArgumentException("A draft needs a stable MessageId so later saves can replace it.", nameof(draft));

        var drafts = await EnsureFolderAsync(account, SpecialFolder.Drafts, "Drafts", ct);
        var mime = BuildMime(account, draft, validateRecipients: false);
        await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
        {
            var folder = await client.GetFolderAsync(drafts.FullName, ct);
            await folder.OpenAsync(FolderAccess.ReadWrite, ct);
            var previous = await folder.SearchAsync(SearchQuery.HeaderContains("Message-ID", draft.MessageId), ct);
            await folder.AppendAsync(new AppendRequest(mime, MessageFlags.Draft | MessageFlags.Seen), ct);
            await ExpungeAsync(client, folder, previous, ct);
            await SyncFolderAsync(folder, drafts.Id, ct);
        }, ct);
    }

    public async Task DeleteDraftAsync(Account account, string messageId, CancellationToken ct = default)
    {
        MailFolder? drafts;
        await using (var db = createContext())
            drafts = await db.Folders.FirstOrDefaultAsync(f => f.AccountId == account.Id && f.Special == SpecialFolder.Drafts, ct);
        if (drafts is null)
            return;

        await ConnectionFor(account, Lane.Interactive).UseAsync(async client =>
        {
            var folder = await client.GetFolderAsync(drafts.FullName, ct);
            await folder.OpenAsync(FolderAccess.ReadWrite, ct);
            var matches = await folder.SearchAsync(SearchQuery.HeaderContains("Message-ID", messageId), ct);
            await ExpungeAsync(client, folder, matches, ct);
            await SyncFolderAsync(folder, drafts.Id, ct);
        }, ct);
    }

    internal static MimeMessage BuildMime(Account account, OutgoingMessage outgoing, bool validateRecipients = true)
    {
        var mime = new MimeMessage();
        mime.From.Add(new MailboxAddress(account.DisplayName, account.Email));
        mime.To.AddRange(ParseRecipients(outgoing.To, validateRecipients));
        mime.Cc.AddRange(ParseRecipients(outgoing.Cc, validateRecipients));
        mime.Bcc.AddRange(ParseRecipients(outgoing.Bcc, validateRecipients));
        if (validateRecipients && mime.To.Count + mime.Cc.Count + mime.Bcc.Count == 0)
            throw new FormatException("Add at least one recipient.");

        mime.Subject = outgoing.Subject;
        mime.MessageId = outgoing.MessageId ?? MimeUtils.GenerateMessageId(AddressList.DomainOf(account.Email));
        if (outgoing.InReplyTo is not null)
            mime.InReplyTo = outgoing.InReplyTo;
        if (outgoing.References is not null)
            foreach (var id in outgoing.References.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                mime.References.Add(id);

        var body = new BodyBuilder { TextBody = outgoing.BodyText, HtmlBody = outgoing.BodyHtml };
        foreach (var a in outgoing.Attachments ?? [])
            body.Attachments.Add(a.FileName, a.Data, ContentType.Parse(a.ContentType));
        mime.Body = outgoing.Calendar is { } calendar ? WithCalendar(body.ToMessageBody(), calendar) : body.ToMessageBody();
        return mime;
    }

    /// <summary>
    /// iMIP layout that Outlook, Gmail and Apple Mail all recognise: the readable text and a
    /// text/calendar part with the iTIP method, side by side in multipart/alternative.
    /// </summary>
    private static MimeEntity WithCalendar(MimeEntity body, OutgoingCalendar calendar)
    {
        var part = new TextPart("calendar") { Text = calendar.Ics };
        part.ContentType.Charset = "utf-8";
        part.ContentType.Parameters.Add("method", calendar.Method);
        part.ContentTransferEncoding = ContentEncoding.QuotedPrintable;

        if (body is Multipart { ContentType.MediaSubtype: "mixed" } mixed)
        {
            // With attachments: put the calendar next to the text inside the first (body) part.
            var alternative = new MultipartAlternative { mixed[0], part };
            mixed.RemoveAt(0);
            mixed.Insert(0, alternative);
            return mixed;
        }
        return body is MultipartAlternative alt ? Append(alt, part) : new MultipartAlternative { body, part };
    }

    private static MultipartAlternative Append(MultipartAlternative alternative, MimeEntity part)
    {
        alternative.Add(part);
        return alternative;
    }

    /// <summary>Drafts keep whatever was typed; sending rejects anything that is not a valid address.</summary>
    private static IEnumerable<InternetAddress> ParseRecipients(IReadOnlyList<string> lines, bool strict)
    {
        foreach (var entry in lines.SelectMany(AddressList.Split))
        {
            if (MailboxAddress.TryParse(entry, out var mailbox) && mailbox.Address.Contains('@'))
                yield return mailbox;
            else if (strict)
                throw new FormatException($"\"{entry}\" isn't a valid email address.");
        }
    }

    // ───────────────────────────── Plumbing ─────────────────────────────

    private async Task<(Account Account, string FolderName)> LocateAsync(int folderId, CancellationToken ct)
    {
        await using var db = createContext();
        var folder = await db.Folders.AsNoTracking().Include(f => f.Account).SingleAsync(f => f.Id == folderId, ct);
        return (folder.Account!, folder.FullName);
    }

    private string PasswordFor(Account account) =>
        credentials.Load(account.Id) ?? throw new InvalidOperationException($"No saved password for {account.Email}.");

    private ImapConnection ConnectionFor(Account account, Lane lane) =>
        _connections.GetOrAdd((account.Id, lane), _ => new ImapConnection(account, () => PasswordFor(account)));

    public async ValueTask DisposeAsync()
    {
        foreach (var c in _connections.Values)
            await c.DisposeAsync();
        _connections.Clear();
    }
}
