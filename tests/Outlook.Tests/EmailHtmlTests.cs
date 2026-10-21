using Outlook.Core.Text;

namespace Outlook.Tests;

public class EmailHtmlTests
{
    [Fact]
    public void Strips_active_content()
    {
        var html = EmailHtml.Sanitize(
            "<p onclick=\"steal()\">Hi</p><script>alert(1)</script><iframe src=\"x\"></iframe>" +
            "<a href=\"javascript:alert(1)\">x</a><form action=\"/\"><input name=\"pw\"></form>");

        Assert.DoesNotContain("onclick", html);
        Assert.DoesNotContain("script", html);
        Assert.DoesNotContain("iframe", html);
        Assert.DoesNotContain("javascript", html);
        Assert.DoesNotContain("<input", html);
        Assert.Contains("<p>Hi</p>", html);
    }

    [Theory]
    [InlineData("<img src=\"https://track.example/p.gif\">", true)]
    [InlineData("<div style=\"background: url('http://x/y.png')\">", true)]
    [InlineData("<link rel=stylesheet href=https://x/s.css>", true)]
    [InlineData("<img src=\"cid:logo@x\">", false)]
    [InlineData("<a href=\"https://example.com\">link</a>", false)]
    public void Detects_remote_content(string html, bool remote)
    {
        Assert.Equal(remote, EmailHtml.Prepare(html, null).HasRemoteContent);
    }

    [Fact]
    public void Plain_text_is_escaped_linked_and_keeps_lines()
    {
        var html = EmailHtml.Prepare(null, "a < b\n  see https://example.com/x.").Html;

        Assert.Contains("a &lt; b<br>", html);
        Assert.Contains("&nbsp;&nbsp;see", html);
        Assert.Contains("<a href=\"https://example.com/x\">https://example.com/x</a>.", html);
    }
}
