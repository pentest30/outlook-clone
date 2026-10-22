using Outlook.Core.Threading;

namespace Outlook.Tests;

public class ThreadResolverTests
{
    [Fact]
    public void Reply_joins_the_thread_of_the_first_reference()
    {
        var root = ThreadResolver.Resolve("<a@x>", null, null, "Offsite");
        var reply = ThreadResolver.Resolve("<b@x>", "<a@x>", "<a@x>", "Re: Offsite");
        var replyToReply = ThreadResolver.Resolve("<c@x>", "<b@x>", "<a@x> <b@x>", "Re: Offsite");

        Assert.Equal("a@x", root);
        Assert.Equal(root, reply);
        Assert.Equal(root, replyToReply);
    }

    [Fact]
    public void Falls_back_to_in_reply_to_when_references_missing()
    {
        Assert.Equal("a@x", ThreadResolver.Resolve("<b@x>", "<A@X>", null, "Re: hi"));
    }

    [Fact]
    public void Without_headers_threads_by_normalized_subject()
    {
        var a = ThreadResolver.Resolve(null, null, null, "Lunch");
        var b = ThreadResolver.Resolve(null, null, null, "RE: Fwd: lunch");
        Assert.Equal(a, b);
    }

    [Theory]
    [InlineData("Re: Re: Plan", "plan")]
    [InlineData("FW: AW: Plan", "plan")]
    [InlineData("Meeting: agenda", "meeting: agenda")]
    [InlineData("  Plan  ", "plan")]
    public void NormalizeSubject_strips_only_reply_and_forward_prefixes(string input, string expected)
    {
        Assert.Equal(expected, ThreadResolver.NormalizeSubject(input));
    }
}
