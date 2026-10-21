// Fills the GreenMail dev server (docker-compose.yml) with a realistic mailbox for dev@localhost:
// standard folders, a nested folder, a reply thread, HTML mail, an attachment, and messages spread
// over the last few weeks so every date group in the list shows up.
//
//   dotnet run --project tools/Outlook.DevSeed [host] [smtpPort] [imapPort]

using MailKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

var host = args.ElementAtOrDefault(0) ?? "localhost";
var smtpPort = int.Parse(args.ElementAtOrDefault(1) ?? "3025");
var imapPort = int.Parse(args.ElementAtOrDefault(2) ?? "3143");
const string me = "dev@localhost";

var people = new[]
{
    ("Ana Lopez", "ana@contoso.test"),
    ("Bilal Haddad", "bilal@fabrikam.test"),
    ("Chen Wei", "chen@northwind.test"),
    ("Dana Smith", "dana@contoso.test"),
    ("GitHub", "noreply@github.test"),
    ("Contoso HR", "hr@contoso.test"),
};
var subjects = new[]
{
    "Quarterly planning notes", "Invoice #4821", "Design review tomorrow", "Your build failed",
    "Lunch on Friday?", "Updated travel policy", "Contract draft v3", "Welcome to the team",
    "Weekly status", "Server maintenance window", "Feedback on the prototype", "Holiday schedule",
};

var now = DateTimeOffset.Now;
var messages = new List<MimeMessage>();
var rng = new Random(42);

for (var i = 0; i < 30; i++)
{
    var (name, address) = people[i % people.Length];
    var m = Build(name, address, subjects[i % subjects.Length],
        $"Hi,\n\nThis is sample message {i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit.\n\nThanks,\n{name}",
        now.AddHours(-(i * i * 1.5 + rng.Next(0, 3))));
    messages.Add(m);
}

// A three-message thread, to exercise References / In-Reply-To threading.
var root = Build("Ana Lopez", "ana@contoso.test", "Offsite venue", "Shall we book the lakeside venue for the offsite?", now.AddHours(-5));
var reply = Build("Bilal Haddad", "bilal@fabrikam.test", "Re: Offsite venue", "Works for me. Can we check catering?", now.AddHours(-4));
reply.InReplyTo = root.MessageId;
reply.References.Add(root.MessageId!);
var reply2 = Build("Ana Lopez", "ana@contoso.test", "Re: Offsite venue", "Catering is included. Booking it now.", now.AddHours(-3));
reply2.InReplyTo = reply.MessageId;
reply2.References.Add(root.MessageId!);
reply2.References.Add(reply.MessageId!);
messages.AddRange([root, reply, reply2]);

// HTML newsletter with an inline logo (cid:) and a remote tracking pixel, which the reading pane blocks.
var html = new MimeMessage { Subject = "This week at Contoso", Date = now.AddDays(-2) };
html.From.Add(new MailboxAddress("Contoso News", "news@contoso.test"));
html.To.Add(MailboxAddress.Parse(me));
var newsletter = new BodyBuilder
{
    HtmlBody = "<html><head><style>p{color:#333} h1{color:#0F6CBD}</style></head><body>" +
               "<img src=\"cid:logo@contoso.test\" alt=\"Contoso\"><h1>This week</h1>" +
               "<p>Our <b>new office</b> opens on Monday.</p><ul><li>Parking: level 2</li><li>Badge pickup: reception</li></ul>" +
               "<p>Read more on <a href=\"https://example.com/news\">the intranet</a> or write to " +
               "<a href=\"mailto:news@contoso.test\">news@contoso.test</a>.</p>" +
               "<p>See you there &amp; welcome!</p><img src=\"https://tracker.example.com/open.gif?id=42\" width=\"1\" height=\"1\">" +
               "</body></html>",
};
var logo = newsletter.LinkedResources.Add("logo.png", Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAKAAAAAoCAIAAAD2TmbPAAAAZUlEQVR42u3YUREAAAREQQ0UVVE2erBmG7wZHxfjTl9kNYcJLDACIzACIzACI/CrwLYeSxZeNAIjMAIjMAILjMBYspwlCy8agRFYYARGYARGYCxZzpKFFy0wAiMwAiMwAiPws8ALAy7mf7H1CNIAAAAASUVORK5CYII="),
    new ContentType("image", "png"));
logo.ContentId = "logo@contoso.test";
html.Body = newsletter.ToMessageBody();
messages.Add(html);

// Message with an attachment, CC'd.
var withFile = Build("Chen Wei", "chen@northwind.test", "Q3 report", "Report attached, let me know what you think.", now.AddMinutes(-40));
withFile.Cc.Add(new MailboxAddress("Dana Smith", "dana@contoso.test"));
var builder = new BodyBuilder { TextBody = "Report attached, let me know what you think." };
builder.Attachments.Add("q3-report.csv", "region,revenue\nnorth,120\nsouth,95\n"u8.ToArray(), new ContentType("text", "csv"));
withFile.Body = builder.ToMessageBody();
messages.Add(withFile);

using (var smtp = new SmtpClient())
{
    await smtp.ConnectAsync(host, smtpPort, SecureSocketOptions.None);
    foreach (var m in messages)
        await smtp.SendAsync(m);
    await smtp.DisconnectAsync(true);
}
Console.WriteLine($"Sent {messages.Count} messages to {me}.");

using (var imap = new ImapClient())
{
    await imap.ConnectAsync(host, imapPort, SecureSocketOptions.None);
    await imap.AuthenticateAsync(me, "dev");
    var personal = imap.GetFolder(imap.PersonalNamespaces[0]);
    var existing = (await imap.GetFoldersAsync(imap.PersonalNamespaces[0])).Select(f => f.FullName).ToHashSet();

    foreach (var name in new[] { "Drafts", "Sent", "Archive", "Junk", "Trash", "Projects" })
        if (!existing.Contains(name))
            await personal.CreateAsync(name, true);
    var projects = await imap.GetFolderAsync("Projects");
    var subfolders = (await projects.GetSubfoldersAsync()).Select(f => f.Name).ToHashSet();
    if (!subfolders.Contains("Alpha"))
        await projects.CreateAsync("Alpha", true);

    // Mark the older half read and flag a couple, so the list shows a realistic mix.
    var inbox = imap.Inbox;
    await inbox.OpenAsync(FolderAccess.ReadWrite);
    var uids = await inbox.SearchAsync(MailKit.Search.SearchQuery.All);
    // Messages arrive newest-first above, so pick the older half by Date header, not by UID.
    var byDate = (await inbox.FetchAsync(uids, new FetchRequest(MessageSummaryItems.UniqueId | MessageSummaryItems.Envelope)))
        .OrderBy(s => s.Envelope?.Date).Select(s => s.UniqueId).ToList();
    var older = byDate.Take(byDate.Count / 2).ToList();
    if (older.Count > 0)
        await inbox.StoreAsync(older, new StoreFlagsRequest(StoreAction.Add, MessageFlags.Seen) { Silent = true });
    if (uids.Count > 3)
        await inbox.StoreAsync([uids[1], uids[^3]], new StoreFlagsRequest(StoreAction.Add, MessageFlags.Flagged) { Silent = true });

    await imap.DisconnectAsync(true);
}
Console.WriteLine("Created folders and set read/flag states.");

static MimeMessage Build(string fromName, string fromAddress, string subject, string body, DateTimeOffset date)
{
    var m = new MimeMessage { Subject = subject, Date = date, MessageId = MimeKit.Utils.MimeUtils.GenerateMessageId("seed.test") };
    m.From.Add(new MailboxAddress(fromName, fromAddress));
    m.To.Add(new MailboxAddress("Dev User", me));
    m.Body = new TextPart("plain") { Text = body };
    return m;
}
