using Outlook.Core.Compose;
using Outlook.Core.Models;
using Outlook.Data;

namespace Outlook.Tests;

public sealed class AddressBookTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-book-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;
    private readonly Account _me;

    public AddressBookTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
        _me = _store.AddAccountAsync(new Account { Email = "me@x.test" }).GetAwaiter().GetResult();
    }

    private async Task Record(bool sent, params (string Name, string Address)[] people)
    {
        await using var db = _store.CreateContext();
        var book = new AddressBook(db, _me);
        book.Record(people, sent, DateTimeOffset.UtcNow);
        await book.SaveAsync(default);
    }

    private async Task<List<string>> Suggest(string prefix) =>
        (await _store.SuggestAddressesAsync(_me.Id, prefix)).Select(a => a.Display).ToList();

    [Fact]
    public async Task Matches_address_first_name_and_last_name()
    {
        await Record(sent: false, ("Ana Lopez", "ana@contoso.test"), ("Bilal Haddad", "bh@fabrikam.test"));

        Assert.Equal(["Ana Lopez <ana@contoso.test>"], await Suggest("an"));
        Assert.Equal(["Ana Lopez <ana@contoso.test>"], await Suggest("lop"));
        Assert.Equal(["Bilal Haddad <bh@fabrikam.test>"], await Suggest("BH"));
        Assert.Empty(await Suggest("x"));
    }

    [Fact]
    public async Task People_I_write_to_rank_above_people_who_only_write_to_me()
    {
        for (var i = 0; i < 4; i++)
            await Record(sent: false, ("Newsletter Team", "news@shop.test"));
        await Record(sent: true, ("Nadia", "nadia@x.test"));

        Assert.Equal(["Nadia <nadia@x.test>", "Newsletter Team <news@shop.test>"], await Suggest("n"));
    }

    [Fact]
    public async Task Skips_my_own_address_and_robots()
    {
        await Record(sent: false, ("Me", "ME@x.test"), ("GitHub", "noreply@github.test"), ("Shop", "do-not-reply@shop.test"),
            ("Mail Delivery", "MAILER-DAEMON@x.test"), ("Real", "real@x.test"));

        Assert.Empty(await Suggest("me"));
        Assert.Empty(await Suggest("git"));
        Assert.Empty(await Suggest("mail"));
        Assert.Single(await Suggest("real"));
    }

    [Fact]
    public async Task Built_once_from_cached_mail()
    {
        await using (var db = _store.CreateContext())
        {
            var inbox = new MailFolder { AccountId = _me.Id, FullName = "INBOX", Special = SpecialFolder.Inbox };
            var sent = new MailFolder { AccountId = _me.Id, FullName = "Sent", Special = SpecialFolder.Sent };
            db.Folders.AddRange(inbox, sent);
            await db.SaveChangesAsync();
            db.Messages.AddRange(
                new MailMessage { FolderId = inbox.Id, Uid = 1, FromName = "Ana Lopez", FromAddress = "ana@contoso.test", ThreadId = "a" },
                new MailMessage { FolderId = sent.Id, Uid = 1, To = "\"Smith, Carl\" <carl@z.test>, dana@contoso.test", ThreadId = "b" });
            await db.SaveChangesAsync();
        }

        await _store.BuildAddressBookAsync(_me);
        await _store.BuildAddressBookAsync(_me);   // second call is a no-op

        Assert.Equal(["Ana Lopez <ana@contoso.test>"], await Suggest("ana"));
        Assert.Equal(["Smith, Carl <carl@z.test>"], await Suggest("carl"));
        var dana = Assert.Single(await _store.SuggestAddressesAsync(_me.Id, "dana"));
        Assert.Equal(1, dana.SentCount);
    }

    [Theory]
    [InlineData("", "")]
    [InlineData("an", "an")]
    [InlineData("ana@x.test; bo", "bo")]
    [InlineData("a@x, \"Smith\" <s@y>;  ca", "ca")]
    public void LastEntry_is_what_is_being_typed(string line, string expected)
    {
        Assert.Equal(expected, AddressList.LastEntry(line));
    }

    [Theory]
    [InlineData("an", "Ana <ana@x>", "Ana <ana@x>; ")]
    [InlineData("bob@y; an", "Ana <ana@x>", "bob@y; Ana <ana@x>; ")]
    [InlineData("bob@y,an", "Ana <ana@x>", "bob@y, Ana <ana@x>; ")]
    public void CompleteLast_replaces_only_the_partial_entry(string line, string picked, string expected)
    {
        Assert.Equal(expected, AddressList.CompleteLast(line, picked));
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
