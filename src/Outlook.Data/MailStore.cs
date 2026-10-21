using System.Linq.Expressions;
using Microsoft.EntityFrameworkCore;
using Outlook.Core.Models;
using Outlook.Core.Search;
using Outlook.Core.Services;

namespace Outlook.Data;

/// <summary>
/// Read side of the local cache. Each call uses a short-lived DbContext so the UI thread and
/// the sync engine never share one (DbContext is not thread-safe).
/// </summary>
public class MailStore(Func<MailDbContext> createContext) : IMailStore
{
    public Func<MailDbContext> CreateContext { get; } = createContext;

    /// <summary>List rows only: bodies stay in the database until a message is opened (GetMessageAsync).</summary>
    private static readonly Expression<Func<MailMessage, MailMessage>> ListRow = m => new MailMessage
    {
        Id = m.Id, FolderId = m.FolderId, Uid = m.Uid,
        MessageIdHeader = m.MessageIdHeader, InReplyTo = m.InReplyTo, References = m.References, ThreadId = m.ThreadId,
        Subject = m.Subject, FromName = m.FromName, FromAddress = m.FromAddress,
        To = m.To, Cc = m.Cc, ReplyTo = m.ReplyTo,
        Date = m.Date, Size = m.Size, Preview = m.Preview,
        IsRead = m.IsRead, IsFlagged = m.IsFlagged, IsAnswered = m.IsAnswered, IsForwarded = m.IsForwarded,
        IsDraft = m.IsDraft, HasAttachments = m.HasAttachments, IsBodyDownloaded = m.IsBodyDownloaded,
        CategoryKeywords = m.CategoryKeywords,
    };

    /// <summary>Creates the cache or upgrades it to the current schema (EF Core migrations).</summary>
    public void Migrate()
    {
        using var db = CreateContext();
        db.Database.Migrate();
    }

    public async Task<IReadOnlyList<Account>> GetAccountsAsync(CancellationToken ct = default)
    {
        await using var db = CreateContext();
        return await db.Accounts.AsNoTracking().OrderBy(a => a.Id).ToListAsync(ct);
    }

    public async Task<Account> AddAccountAsync(Account account, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        db.Accounts.Add(account);
        await db.SaveChangesAsync(ct);
        return account;
    }

    public async Task<IReadOnlyList<MailFolder>> GetFoldersAsync(int accountId, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        var folders = await db.Folders.AsNoTracking().Where(f => f.AccountId == accountId).ToListAsync(ct);
        return folders.OrderBy(f => f.SortRank).ThenBy(f => f.FullName, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public async Task<IReadOnlyList<MailMessage>> GetMessagesAsync(int folderId, int take = 500, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        return await db.Messages.AsNoTracking()
            .Where(m => m.FolderId == folderId)
            .OrderByDescending(m => m.Date)
            .Take(take)
            .Select(ListRow)
            .ToListAsync(ct);
    }

    public async Task<MailMessage?> GetMessageAsync(int messageId, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        return await db.Messages.AsNoTracking().SingleOrDefaultAsync(m => m.Id == messageId, ct);
    }

    public async Task<IReadOnlyList<MailMessage>> SearchAsync(int accountId, int? folderId, MailSearchQuery query, int take = 300,
        CancellationToken ct = default)
    {
        await using var db = CreateContext();

        // Text parts go through the FTS5 index (see the AddSearchAndCategories migration); the rest are plain filters.
        var match = query.ToFtsMatch();
        var messages = match is null
            ? db.Messages.AsNoTracking()
            : db.Messages.FromSql($"SELECT * FROM Messages WHERE Id IN (SELECT rowid FROM MessageSearch WHERE MessageSearch MATCH {match})").AsNoTracking();

        messages = messages.Where(m => m.Folder!.AccountId == accountId);
        if (folderId is not null)
            messages = messages.Where(m => m.FolderId == folderId);
        if (query.IsRead is { } isRead)
            messages = messages.Where(m => m.IsRead == isRead);
        if (query.IsFlagged is { } isFlagged)
            messages = messages.Where(m => m.IsFlagged == isFlagged);
        if (query.HasAttachment is { } hasAttachment)
            messages = messages.Where(m => m.HasAttachments == hasAttachment);

        if (query.Categories.Count > 0)
        {
            // "category:red" matches any category whose name contains "red".
            var categories = await db.Categories.AsNoTracking().Where(c => c.AccountId == accountId).ToListAsync(ct);
            foreach (var wanted in query.Categories)
            {
                var keywords = categories.Where(c => c.Name.Contains(wanted, StringComparison.OrdinalIgnoreCase))
                    .Select(c => c.Keyword).DefaultIfEmpty(Category.KeywordFor(wanted)).ToList();
                messages = messages.Where(HasAnyKeyword(keywords));
            }
        }

        return await messages.OrderByDescending(m => m.Date).Take(take).Select(ListRow).ToListAsync(ct);
    }

    /// <summary>
    /// m => " kw1 kw2 ".Contains(" a ") || " kw1 kw2 ".Contains(" b ") ..., padded so only whole keywords
    /// match. Built by hand because EF cannot translate Any() over a local list inside this predicate.
    /// </summary>
    private static Expression<Func<MailMessage, bool>> HasAnyKeyword(IEnumerable<string> keywords)
    {
        var m = Expression.Parameter(typeof(MailMessage), "m");
        var concat = typeof(string).GetMethod(nameof(string.Concat), [typeof(string), typeof(string), typeof(string)])!;
        var contains = typeof(string).GetMethod(nameof(string.Contains), [typeof(string)])!;
        var padded = Expression.Call(concat, Expression.Constant(" "),
            Expression.Property(m, nameof(MailMessage.CategoryKeywords)), Expression.Constant(" "));

        Expression? body = null;
        foreach (var keyword in keywords)
        {
            var test = Expression.Call(padded, contains, Expression.Constant($" {keyword} "));
            body = body is null ? test : Expression.OrElse(body, test);
        }
        return Expression.Lambda<Func<MailMessage, bool>>(body ?? Expression.Constant(false), m);
    }

    public async Task<IReadOnlyList<MailMessage>> GetThreadMessagesAsync(int accountId, IReadOnlyCollection<string> threadIds,
        CancellationToken ct = default)
    {
        if (threadIds.Count == 0) return [];
        await using var db = CreateContext();
        var result = new List<MailMessage>();
        // SQLite caps the number of query parameters, so ask in chunks.
        foreach (var chunk in threadIds.Chunk(500))
        {
            result.AddRange(await db.Messages.AsNoTracking()
                .Where(m => m.Folder!.AccountId == accountId && chunk.Contains(m.ThreadId))
                .Select(ListRow)
                .ToListAsync(ct));
        }
        return result;
    }

    public async Task<IReadOnlyList<Category>> GetCategoriesAsync(int accountId, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        var categories = await db.Categories.AsNoTracking().Where(c => c.AccountId == accountId).OrderBy(c => c.Id).ToListAsync(ct);
        if (categories.Count > 0)
            return categories;

        var defaults = Category.Defaults
            .Select(d => new Category { AccountId = accountId, Name = d.Name, Color = d.Color, Keyword = Category.KeywordFor(d.Name) })
            .ToList();
        db.Categories.AddRange(defaults);
        await db.SaveChangesAsync(ct);
        return defaults;
    }

    public async Task<Category> AddCategoryAsync(int accountId, string name, string color, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        var keyword = Category.KeywordFor(name);
        var existing = await db.Categories.FirstOrDefaultAsync(c => c.AccountId == accountId && c.Keyword == keyword, ct);
        if (existing is not null)
            return existing;
        var category = new Category { AccountId = accountId, Name = name.Trim(), Color = color, Keyword = keyword };
        db.Categories.Add(category);
        await db.SaveChangesAsync(ct);
        return category;
    }

    public async Task<IReadOnlyList<MailAttachment>> GetAttachmentsAsync(int messageId, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        // Project so the (possibly large) Data column is never read here.
        return await db.Attachments.AsNoTracking()
            .Where(a => a.MessageId == messageId)
            .OrderBy(a => a.Id)
            .Select(a => new MailAttachment
            {
                Id = a.Id, MessageId = a.MessageId, FileName = a.FileName, ContentType = a.ContentType,
                Size = a.Size, ContentId = a.ContentId, IsInline = a.IsInline,
            })
            .ToListAsync(ct);
    }

    public async Task<byte[]> GetAttachmentDataAsync(int attachmentId, CancellationToken ct = default)
    {
        await using var db = CreateContext();
        return await db.Attachments.Where(a => a.Id == attachmentId).Select(a => a.Data).SingleAsync(ct);
    }
}
