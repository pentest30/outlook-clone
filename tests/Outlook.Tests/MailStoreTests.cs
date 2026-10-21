using Outlook.Core.Models;
using Outlook.Data;

namespace Outlook.Tests;

public sealed class MailStoreTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-test-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;

    public MailStoreTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
    }

    [Fact]
    public async Task Folders_come_back_in_outlook_order()
    {
        var account = await _store.AddAccountAsync(new Account { Email = "me@x.test" });
        await using (var db = _store.CreateContext())
        {
            db.Folders.AddRange(
                new MailFolder { AccountId = account.Id, FullName = "Zebra", DisplayName = "Zebra" },
                new MailFolder { AccountId = account.Id, FullName = "Trash", DisplayName = "Trash", Special = SpecialFolder.Trash },
                new MailFolder { AccountId = account.Id, FullName = "Apple", DisplayName = "Apple" },
                new MailFolder { AccountId = account.Id, FullName = "Sent", DisplayName = "Sent", Special = SpecialFolder.Sent },
                new MailFolder { AccountId = account.Id, FullName = "INBOX", DisplayName = "Inbox", Special = SpecialFolder.Inbox });
            await db.SaveChangesAsync();
        }

        var names = (await _store.GetFoldersAsync(account.Id)).Select(f => f.DisplayName);

        Assert.Equal(["Inbox", "Sent", "Trash", "Apple", "Zebra"], names);
    }

    [Fact]
    public async Task Messages_are_newest_first_and_keep_their_offset_exact()
    {
        var account = await _store.AddAccountAsync(new Account { Email = "me@x.test" });
        var older = new DateTimeOffset(2026, 3, 1, 9, 0, 0, TimeSpan.FromHours(2));
        var newer = new DateTimeOffset(2026, 3, 1, 8, 30, 0, TimeSpan.Zero);   // later instant despite the smaller clock time
        int folderId;
        await using (var db = _store.CreateContext())
        {
            var folder = new MailFolder { AccountId = account.Id, FullName = "INBOX" };
            db.Folders.Add(folder);
            await db.SaveChangesAsync();
            folderId = folder.Id;
            db.Messages.AddRange(
                new MailMessage { FolderId = folderId, Uid = 1, Subject = "older", Date = older },
                new MailMessage { FolderId = folderId, Uid = 2, Subject = "newer", Date = newer });
            await db.SaveChangesAsync();
        }

        var messages = await _store.GetMessagesAsync(folderId);

        Assert.Equal(["newer", "older"], messages.Select(m => m.Subject));
        Assert.Equal(older.UtcDateTime, messages[1].Date.UtcDateTime);
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
