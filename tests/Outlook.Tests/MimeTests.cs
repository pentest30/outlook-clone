using MimeKit;
using Outlook.Core.Models;
using Outlook.Mail;

namespace Outlook.Tests;

public class MimeTests
{
    private static readonly Account Me = new() { DisplayName = "Me", Email = "me@x.test" };

    [Fact]
    public void Extracts_files_and_inline_pictures_but_not_the_body()
    {
        var builder = new BodyBuilder { TextBody = "text", HtmlBody = "<img src=\"cid:logo@x\">" };
        var logo = builder.LinkedResources.Add("logo.png", [1, 2, 3], new ContentType("image", "png"));
        logo.ContentId = "logo@x";
        builder.Attachments.Add("..\\..\\evil/report.pdf", [4, 5], new ContentType("application", "pdf"));
        var mime = new MimeMessage { Body = builder.ToMessageBody() };

        var parts = ImapMailService.ExtractAttachments(mime);

        Assert.Equal(2, parts.Count);
        var file = Assert.Single(parts, p => !p.IsInline);
        Assert.Equal("report.pdf", file.FileName);   // path parts from the sender are dropped
        Assert.Equal([4, 5], file.Data);
        var inline = Assert.Single(parts, p => p.IsInline);
        Assert.Equal("logo@x", inline.ContentId);
        Assert.Equal(3, inline.Size);
    }

    [Fact]
    public void Builds_recipients_threading_and_attachments()
    {
        var mime = ImapMailService.BuildMime(Me, new OutgoingMessage(
            ["\"Smith, Bob\" <bob@y.test>; ana@z.test"], [], [], "RE: Hi", "Body",
            InReplyTo: "orig@x", References: "root@x orig@x",
            Attachments: [new OutgoingAttachment("a.txt", "text/plain", "hello"u8.ToArray())],
            MessageId: "fixed@x.test"));

        Assert.Equal(["bob@y.test", "ana@z.test"], mime.To.Mailboxes.Select(m => m.Address));
        Assert.Equal("Smith, Bob", mime.To.Mailboxes.First().Name);
        Assert.Equal("orig@x", mime.InReplyTo);
        Assert.Equal(["root@x", "orig@x"], mime.References);
        Assert.Equal("fixed@x.test", mime.MessageId);
        Assert.Equal("a.txt", Assert.Single(mime.Attachments).ContentDisposition?.FileName);
    }

    [Fact]
    public void Addresses_display_unquoted_unless_quoting_is_needed()
    {
        var list = InternetAddressList.Parse("\"Dev User\" <dev@x>, \"Smith, Bob\" <bob@y>, carl@z");

        Assert.Equal("Dev User <dev@x>, \"Smith, Bob\" <bob@y>, carl@z", ImapMailService.FormatAddresses(list));
    }

    [Fact]
    public void Sending_rejects_bad_addresses_but_drafts_keep_going()
    {
        var message = new OutgoingMessage(["not an address"], [], [], "s", "b", MessageId: "d@x");

        Assert.Throws<FormatException>(() => ImapMailService.BuildMime(Me, message));
        Assert.Empty(ImapMailService.BuildMime(Me, message, validateRecipients: false).To);
    }
}
