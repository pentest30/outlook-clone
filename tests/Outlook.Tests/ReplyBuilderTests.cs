using Outlook.Core.Compose;
using Outlook.Core.Models;

namespace Outlook.Tests;

public class ReplyBuilderTests
{
    private static MailMessage Original() => new()
    {
        MessageIdHeader = "orig@x",
        References = "root@x",
        Subject = "Budget",
        FromName = "Ana Lopez",
        FromAddress = "ana@contoso.test",
        To = "me@x.test, Bob <bob@y.test>",
        Cc = "\"Smith, Carl\" <carl@z.test>, ME@x.test, bob@y.test",
        Date = new DateTimeOffset(2026, 10, 1, 9, 30, 0, TimeSpan.Zero),
    };

    [Fact]
    public void Reply_goes_to_sender_with_threading_headers()
    {
        var seed = ReplyBuilder.Create(ComposeMode.Reply, Original(), "Original text", "me@x.test");

        Assert.Equal("Ana Lopez <ana@contoso.test>", seed.To);
        Assert.Equal("", seed.Cc);
        Assert.Equal("RE: Budget", seed.Subject);
        Assert.Equal("orig@x", seed.InReplyTo);
        Assert.Equal("root@x orig@x", seed.References);
        Assert.Contains(ReplyBuilder.Separator, seed.Body);
        Assert.Contains("From: Ana Lopez <ana@contoso.test>", seed.Body);
        Assert.EndsWith("Original text", seed.Body);
    }

    [Fact]
    public void Reply_prefers_reply_to_header()
    {
        var original = Original();
        original.ReplyTo = "list@contoso.test";

        Assert.Equal("list@contoso.test", ReplyBuilder.Create(ComposeMode.Reply, original, "", "me@x.test").To);
    }

    [Fact]
    public void Reply_all_drops_me_and_duplicates()
    {
        var seed = ReplyBuilder.Create(ComposeMode.ReplyAll, Original(), "", "me@x.test");

        Assert.Equal(["Ana Lopez <ana@contoso.test>", "Bob <bob@y.test>"], AddressList.Split(seed.To));
        // bob is already on To, and "ME@x.test" is me in another case.
        Assert.Equal(["\"Smith, Carl\" <carl@z.test>"], AddressList.Split(seed.Cc));
    }

    [Fact]
    public void Forward_has_no_recipients_and_no_in_reply_to()
    {
        var seed = ReplyBuilder.Create(ComposeMode.Forward, Original(), "", "me@x.test");

        Assert.Equal("", seed.To);
        Assert.Equal("FW: Budget", seed.Subject);
        Assert.Null(seed.InReplyTo);
        Assert.Equal("root@x orig@x", seed.References);
    }

    [Theory]
    [InlineData("RE:", "Re: Budget", "Re: Budget")]
    [InlineData("RE:", "Budget", "RE: Budget")]
    [InlineData("FW:", "RE: Budget", "FW: RE: Budget")]
    public void Prefix_is_not_doubled(string prefix, string subject, string expected)
    {
        Assert.Equal(expected, ReplyBuilder.Prefix(prefix, subject));
    }
}

public class AddressListTests
{
    [Fact]
    public void Splits_on_commas_and_semicolons_outside_quotes_and_brackets()
    {
        var entries = AddressList.Split("\"Smith, Bob\" <bob@y>; ana@x,  <odd,addr@z> ;");

        Assert.Equal(["\"Smith, Bob\" <bob@y>", "ana@x", "<odd,addr@z>"], entries);
    }

    [Theory]
    [InlineData("Ana <ANA@x.test>", "ana@x.test")]
    [InlineData(" bob@y.test ", "bob@y.test")]
    public void AddressOf_extracts_lowercase_address(string entry, string expected)
    {
        Assert.Equal(expected, AddressList.AddressOf(entry));
    }
}
