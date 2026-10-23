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

    [Fact]
    public void Drops_images_without_a_source()
    {
        Assert.Equal("<p>a</p><img src=\"cid:x\">", EmailHtml.Sanitize("<p>a</p><img class=\"image\" src=\"\"><img src=\"cid:x\">"));
    }

    [Fact]
    public void Text_cells_lose_their_fixed_height_but_spacers_keep_it()
    {
        var html = "<table><tr><td style=\"height: 20px\">&nbsp;</td></tr>" +
                   "<tr><td height=\"92\" style=\"height: 92px; font-size: 34px; line-height: 46px;\"><b>A long title</b></td></tr></table>";
        Assert.Equal("<table><tr><td style=\"height: 20px\">&nbsp;</td></tr>" +
                     "<tr><td style=\"font-size: 34px; line-height: 46px;\"><b>A long title</b></td></tr></table>",
            EmailHtml.Sanitize(html));
    }

    [Theory]
    [InlineData("<td style=\"background: #0F6CBD; color: white\">", "<td style=\"background: #0F6CBD; background-color: #0F6CBD; color: white\">")]
    [InlineData("<td style=\"background:white\">", "<td style=\"background:white; background-color: white\">")]
    [InlineData("<style>.h { background: rgb(1, 2, 3) url(x.png) no-repeat; }</style>",
                "<style>.h { background: rgb(1, 2, 3) url(x.png) no-repeat; background-color: rgb(1, 2, 3); }</style>")]
    [InlineData("<td style=\"background-color: red\">", "<td style=\"background-color: red\">")]
    [InlineData("<td style=\"background: none\">", "<td style=\"background: none\">")]
    public void Background_shorthand_colours_are_spelled_out(string html, string expected)
    {
        Assert.Equal(expected, EmailHtml.Sanitize(html));
    }

    [Theory]
    [InlineData("<ul style=\"line-height: 1.4; list-style-type: square; padding-left: 15px\">",
                "<ul style=\"line-height: 1.4; list-style-type: disc; padding-left: 15px\">")]
    [InlineData("<style>ul { list-style: square inside; }</style>", "<style>ul { list-style: disc inside; }</style>")]
    [InlineData("<p class=\"square\">square</p>", "<p class=\"square\">square</p>")]
    public void Square_bullets_become_round(string html, string expected)
    {
        Assert.Equal(expected, EmailHtml.Sanitize(html));
    }

    [Fact]
    public void Blocked_web_pictures_give_way_to_their_alt_text()
    {
        var html = "<p><img src=\"https://x.test/logo.png\" alt=\"Render\" style=\"width: 100%; height: auto\"></p>" +
                   "<img src='https://track.test/open?m=1' width=1 height=1>" +
                   "<img src=\"cid:inline@x\" alt=\"inline\">";
        Assert.Equal("<p><span style=\"color: #616161\">Render</span></p><img src=\"cid:inline@x\" alt=\"inline\">",
            EmailHtml.WithoutRemoteImages(html));
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
