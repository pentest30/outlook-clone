using Outlook.Core.Text;

namespace Outlook.Tests;

public class HtmlTextTests
{
    [Fact]
    public void Strips_tags_styles_and_decodes_entities()
    {
        var text = HtmlText.ToPlainText(
            "<html><head><style>p{color:red}</style></head><body><p>Hello &amp; welcome</p><p>Line&nbsp;two<br>three</p></body></html>");

        Assert.Equal("Hello & welcome\n\nLine two\nthree", text);
    }

    [Fact]
    public void Drops_scripts()
    {
        Assert.Equal("Safe", HtmlText.ToPlainText("<script>alert(1)</script>Safe"));
    }
}
