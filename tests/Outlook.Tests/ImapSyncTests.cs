using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Data;
using Outlook.Mail;

namespace Outlook.Tests;

/// <summary>
/// End-to-end sync against the GreenMail container. Skipped unless GREENMAIL=1, e.g.
///   docker compose up -d; $env:GREENMAIL=1; dotnet test
/// </summary>
public sealed class ImapSyncTests : IAsyncLifetime
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-sync-{Guid.NewGuid():N}.db");
    private MailStore _store = null!;
    private ImapMailService _service = null!;
    private Account _account = null!;

    private static bool Enabled => Environment.GetEnvironmentVariable("GREENMAIL") == "1";

    public async Task InitializeAsync()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
        var credentials = new InMemoryCredentialStore();
        _service = new ImapMailService(_store.CreateContext, credentials);

        // A fresh mailbox per run: GreenMail creates the user on first login.
        var email = $"sync-{Guid.NewGuid():N}@localhost";
        _account = await _store.AddAccountAsync(new Account
        {
            DisplayName = "Sync Test", Email = email, UserName = email,
            ImapHost = "localhost", ImapPort = 3143, ImapSecurity = TransportSecurity.None,
            SmtpHost = "localhost", SmtpPort = 3025, SmtpSecurity = TransportSecurity.None,
        });
        credentials.Save(_account.Id, "pw");
    }

    [Fact]
    public async Task Sent_mail_arrives_syncs_and_round_trips_flags()
    {
        if (!Enabled) return;

        await _service.VerifyAsync(_account, "pw");
        await _service.SendAsync(_account, new OutgoingMessage([_account.Email], [], [], "Hello sync", "Body text"));
        await _service.SyncAccountAsync(_account);

        var inbox = (await _store.GetFoldersAsync(_account.Id)).Single(f => f.Special == SpecialFolder.Inbox);
        var message = Assert.Single(await _store.GetMessagesAsync(inbox.Id));
        Assert.Equal("Hello sync", message.Subject);
        Assert.False(message.IsRead);
        Assert.Equal(1, inbox.UnreadCount);

        var full = await _service.DownloadBodyAsync(message);
        Assert.Contains("Body text", full.BodyText);

        await _service.SetReadAsync(full, true);
        await _service.SetFlaggedAsync(full, true);

        // Wipe local state and resync: the flags must now come from the server.
        await using (var db = _store.CreateContext())
        {
            db.Messages.RemoveRange(db.Messages);
            await db.SaveChangesAsync();
        }
        await _service.SyncAccountAsync(_account);
        var resynced = Assert.Single(await _store.GetMessagesAsync(inbox.Id));
        Assert.True(resynced.IsRead);
        Assert.True(resynced.IsFlagged);
    }

    [Fact]
    public async Task Attachments_are_cached_when_the_body_downloads()
    {
        if (!Enabled) return;

        await SendToSelfAsync("With file", [new OutgoingAttachment("notes.txt", "text/plain", "hello"u8.ToArray())]);
        var message = Assert.Single(await InboxMessagesAsync());
        Assert.True(message.HasAttachments);

        await _service.DownloadBodyAsync(message);

        var attachment = Assert.Single(await _store.GetAttachmentsAsync(message.Id));
        Assert.Equal("notes.txt", attachment.FileName);
        Assert.Empty(attachment.Data);   // metadata query leaves the bytes out
        Assert.Equal("hello"u8.ToArray(), await _store.GetAttachmentDataAsync(attachment.Id));
    }

    [Fact]
    public async Task Archive_creates_the_folder_and_moves_the_message()
    {
        if (!Enabled) return;

        await SendToSelfAsync("Keep me");
        var message = Assert.Single(await InboxMessagesAsync());

        await _service.ArchiveAsync([message]);
        await _service.SyncAccountAsync(_account);

        Assert.Empty(await InboxMessagesAsync());
        var archive = (await _store.GetFoldersAsync(_account.Id)).Single(f => f.Special == SpecialFolder.Archive);
        Assert.Equal("Keep me", Assert.Single(await _store.GetMessagesAsync(archive.Id)).Subject);
    }

    [Fact]
    public async Task Delete_without_a_trash_folder_removes_from_the_server()
    {
        if (!Enabled) return;

        await SendToSelfAsync("Bye");
        var message = Assert.Single(await InboxMessagesAsync());

        await _service.DeleteAsync([message]);
        Assert.Empty(await InboxMessagesAsync());

        // Not just the cache: a full resync must not bring it back.
        await _service.SyncAccountAsync(_account);
        Assert.Empty(await InboxMessagesAsync());
    }

    [Fact]
    public async Task Saving_a_draft_twice_keeps_one_copy_and_delete_removes_it()
    {
        if (!Enabled) return;

        var draft = new OutgoingMessage(["someone@localhost"], [], [], "Plan", "v1", MessageId: $"{Guid.NewGuid():N}@localhost");
        await _service.SaveDraftAsync(_account, draft);
        await _service.SaveDraftAsync(_account, draft with { BodyText = "v2" });

        var drafts = (await _store.GetFoldersAsync(_account.Id)).Single(f => f.Special == SpecialFolder.Drafts);
        var saved = Assert.Single(await _store.GetMessagesAsync(drafts.Id));
        Assert.True(saved.IsDraft);
        Assert.Contains("v2", (await _service.DownloadBodyAsync(saved)).BodyText);

        await _service.DeleteDraftAsync(_account, draft.MessageId!);
        Assert.Empty(await _store.GetMessagesAsync(drafts.Id));
    }

    [Fact]
    public async Task Reply_marks_the_original_answered()
    {
        if (!Enabled) return;

        await SendToSelfAsync("Question");
        var message = Assert.Single(await InboxMessagesAsync());

        await _service.MarkRespondedAsync(message, forwarded: false);
        await _service.MarkRespondedAsync(message, forwarded: true);
        await using (var db = _store.CreateContext())
        {
            db.Messages.RemoveRange(db.Messages);
            await db.SaveChangesAsync();
        }
        await _service.SyncAccountAsync(_account);

        var resynced = Assert.Single(await InboxMessagesAsync());
        Assert.True(resynced.IsAnswered);
        Assert.True(resynced.IsForwarded);
    }

    [Fact]
    public async Task Push_reports_new_inbox_mail()
    {
        if (!Enabled) return;

        await _service.SyncAccountAsync(_account);
        var changed = new TaskCompletionSource<FolderChangedEventArgs>(TaskCreationOptions.RunContinuationsAsynchronously);
        _service.FolderChanged += (_, e) => changed.TrySetResult(e);
        using var cts = new CancellationTokenSource();
        var push = _service.RunPushAsync(_account, cts.Token);

        await Task.Delay(1000);   // let IDLE start
        Assert.False(push.IsCompleted, "Server has no IDLE support, so push returned immediately.");
        await _service.SendAsync(_account, new OutgoingMessage([_account.Email], [], [], "Ping", "x"));

        var e = await changed.Task.WaitAsync(TimeSpan.FromSeconds(15));
        Assert.Equal(_account.Id, e.AccountId);
        Assert.Equal("INBOX", e.FolderFullName);
        await cts.CancelAsync();
        await push.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public async Task Categories_are_stored_on_the_server()
    {
        if (!Enabled) return;

        await SendToSelfAsync("Tag me");
        var message = Assert.Single(await InboxMessagesAsync());
        var red = Category.KeywordFor("Red category");

        await _service.SetCategoryAsync([message], red, add: true);
        await using (var db = _store.CreateContext())
        {
            db.Messages.RemoveRange(db.Messages);
            await db.SaveChangesAsync();
        }
        await _service.SyncAccountAsync(_account);
        var tagged = Assert.Single(await InboxMessagesAsync());
        Assert.Equal(red, tagged.CategoryKeywords);

        await _service.SetCategoryAsync([tagged], red, add: false);
        await _service.SyncAccountAsync(_account);
        Assert.Equal("", Assert.Single(await InboxMessagesAsync()).CategoryKeywords);
    }

    [Fact]
    public async Task Interactive_actions_reconnect_after_the_session_is_dropped()
    {
        if (!Enabled) return;

        await using var proxy = new DroppingProxy("localhost", 3143);
        _account.ImapPort = proxy.Port;
        await SendToSelfAsync("Still there?");
        var message = Assert.Single(await InboxMessagesAsync());
        await _service.SetReadAsync(message, true);   // opens the interactive session

        proxy.DropAll();   // the server (or a router) silently kills idle sessions

        await _service.SetFlaggedAsync(message, true);
        await _service.SyncAccountAsync(_account);
        Assert.True(Assert.Single(await InboxMessagesAsync()).IsFlagged);
    }

    [Fact]
    public async Task A_failed_copy_to_sent_does_not_report_the_send_as_failed()
    {
        if (!Enabled) return;

        await using var proxy = new DroppingProxy("localhost", 3143);
        _account.ImapPort = proxy.Port;
        await _service.SyncAccountAsync(_account);
        await using (var db = _store.CreateContext())
        {
            // Pretend the account has a Sent folder the server does not actually have, so APPEND fails.
            db.Folders.Add(new MailFolder { AccountId = _account.Id, FullName = "Sent-missing", DisplayName = "Sent", Special = SpecialFolder.Sent });
            await db.SaveChangesAsync();
        }

        var report = await _service.SendAsync(_account, new OutgoingMessage([_account.Email], [], [], "Delivered", "x"));

        Assert.False(report.SavedToSent);
        Assert.StartsWith("Message sent", report.Problem);
        _account.ImapPort = 3143;
        await _service.SyncAccountAsync(_account);
        Assert.Contains(await InboxMessagesAsync(), m => m.Subject == "Delivered");
    }

    private async Task SendToSelfAsync(string subject, IReadOnlyList<OutgoingAttachment>? attachments = null)
    {
        await _service.SendAsync(_account, new OutgoingMessage([_account.Email], [], [], subject, "Body", Attachments: attachments));
        await _service.SyncAccountAsync(_account);
    }

    private async Task<IReadOnlyList<MailMessage>> InboxMessagesAsync()
    {
        var inbox = (await _store.GetFoldersAsync(_account.Id)).Single(f => f.Special == SpecialFolder.Inbox);
        return await _store.GetMessagesAsync(inbox.Id);
    }

    public async Task DisposeAsync()
    {
        await _service.DisposeAsync();
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}
