using Outlook.Core.Models;
using Outlook.Core.Search;
using Outlook.Data;

namespace Outlook.Tests;

public class MailSearchQueryTests
{
    [Fact]
    public void Parses_operators_quotes_and_free_words()
    {
        var q = MailSearchQuery.Parse("from:\"Ana Lopez\" budget is:unread has:attachment subject:Q3 category:red \"exact phrase\"");

        Assert.Equal(["Ana Lopez"], q.From);
        Assert.Equal(["Q3"], q.Subject);
        Assert.Equal(["red"], q.Categories);
        Assert.Equal(["budget", "exact phrase"], q.Words);
        Assert.False(q.IsRead);
        Assert.True(q.HasAttachment);
        Assert.Null(q.IsFlagged);
    }

    [Fact]
    public void Unknown_operators_are_plain_words()
    {
        Assert.Equal(["size:big", "http://x"], MailSearchQuery.Parse("size:big http://x").Words);
    }

    [Fact]
    public void Fts_terms_are_quoted_so_input_cannot_inject_syntax()
    {
        var match = MailSearchQuery.Parse("from:ana NOT \"a\"\"b\" x*").ToFtsMatch();

        // Quote characters never survive tokenising, and FTS keywords like NOT or * become plain quoted terms.
        Assert.Equal("\"NOT\"* AND \"ab\"* AND \"x*\"* AND {FromName FromAddress} : \"ana\"*", match);
    }

    [Fact]
    public void Flag_only_queries_skip_the_text_index()
    {
        Assert.Null(MailSearchQuery.Parse("is:flagged").ToFtsMatch());
        Assert.True(MailSearchQuery.Parse("").IsEmpty);
    }
}

public sealed class StoreSearchTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-search-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;
    private readonly Account _account;
    private readonly MailFolder _inbox, _sent;

    public StoreSearchTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
        _account = _store.AddAccountAsync(new Account { Email = "me@x.test" }).GetAwaiter().GetResult();
        using var db = _store.CreateContext();
        _inbox = new MailFolder { AccountId = _account.Id, FullName = "INBOX", Special = SpecialFolder.Inbox };
        _sent = new MailFolder { AccountId = _account.Id, FullName = "Sent", Special = SpecialFolder.Sent };
        db.Folders.AddRange(_inbox, _sent);
        db.SaveChanges();
        db.Messages.AddRange(
            Msg(_inbox, 1, "Contract draft v3", "Ana Lopez", "ana@contoso.test", "t1", read: false, body: "Please review the indemnity clause"),
            Msg(_inbox, 2, "Lunch on Friday?", "Bilal Haddad", "bilal@fabrikam.test", "t2", read: true, keywords: "$Cat_Red_category"),
            Msg(_sent, 3, "RE: Contract draft v3", "Me", "me@x.test", "t1", read: true, attachments: true));
        db.SaveChanges();
    }

    private static MailMessage Msg(MailFolder f, uint uid, string subject, string name, string address, string thread,
        bool read, string? body = null, string keywords = "", bool attachments = false) => new()
    {
        FolderId = f.Id, Uid = uid, Subject = subject, FromName = name, FromAddress = address, ThreadId = thread,
        IsRead = read, SearchText = body, CategoryKeywords = keywords, HasAttachments = attachments,
        Date = DateTimeOffset.UtcNow.AddHours(-uid),
    };

    private Task<IReadOnlyList<MailMessage>> Search(string text, int? folderId = null) =>
        _store.SearchAsync(_account.Id, folderId, MailSearchQuery.Parse(text));

    [Fact]
    public async Task Finds_by_prefix_across_folders_newest_first()
    {
        var results = await Search("contr");
        Assert.Equal(["Contract draft v3", "RE: Contract draft v3"], results.Select(m => m.Subject));
    }

    [Fact]
    public async Task Searches_downloaded_body_text()
    {
        Assert.Equal("Contract draft v3", Assert.Single(await Search("indemnity")).Subject);
    }

    [Fact]
    public async Task Combines_operators_and_scopes_to_a_folder()
    {
        Assert.Equal("Contract draft v3", Assert.Single(await Search("from:ana is:unread")).Subject);
        Assert.Equal("RE: Contract draft v3", Assert.Single(await Search("has:attachment")).Subject);
        Assert.Empty(await Search("contract", _inbox.Id + 1000));
        Assert.Single(await Search("contract", _sent.Id));
    }

    [Fact]
    public async Task Category_filter_matches_by_name()
    {
        await _store.GetCategoriesAsync(_account.Id);   // seeds "Red category" → $Cat_Red_category
        Assert.Equal("Lunch on Friday?", Assert.Single(await Search("category:red")).Subject);
    }

    [Fact]
    public async Task Index_follows_updates_and_deletes()
    {
        await using (var db = _store.CreateContext())
        {
            var lunch = db.Messages.Single(m => m.Uid == 2);
            lunch.Subject = "Dinner on Friday?";
            db.Messages.Remove(db.Messages.Single(m => m.Uid == 1));
            await db.SaveChangesAsync();
        }

        Assert.Empty(await Search("lunch"));
        Assert.Single(await Search("dinner"));
        Assert.Equal("RE: Contract draft v3", Assert.Single(await Search("contract")).Subject);
    }

    [Fact]
    public async Task Thread_messages_span_folders()
    {
        var thread = await _store.GetThreadMessagesAsync(_account.Id, ["t1"]);
        Assert.Equal(2, thread.Count);
        Assert.Contains(thread, m => m.FolderId == _sent.Id);
    }

    [Fact]
    public async Task Default_categories_are_created_once()
    {
        var first = await _store.GetCategoriesAsync(_account.Id);
        var second = await _store.GetCategoriesAsync(_account.Id);
        Assert.Equal(6, first.Count);
        Assert.Equal(first.Select(c => c.Id), second.Select(c => c.Id));
        Assert.Equal("$Cat_Purple_category", second[^1].Keyword);
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
