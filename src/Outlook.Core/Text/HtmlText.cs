using System.Net;
using System.Text.RegularExpressions;

namespace Outlook.Core.Text;

/// <summary>
/// Turns an HTML body into readable plain text for the Phase 0 reading pane, which cannot render HTML yet.
/// </summary>
public static partial class HtmlText
{
    public static string ToPlainText(string html)
    {
        var s = HiddenBlocks().Replace(html, "");
        s = LineBreaks().Replace(s, "\n");
        s = BlockEnds().Replace(s, "\n\n");
        s = Tags().Replace(s, "");
        s = WebUtility.HtmlDecode(s);
        s = SpaceRuns().Replace(s, " ");
        s = BlankLines().Replace(s, "\n\n");
        return s.Trim();
    }

    [GeneratedRegex(@"<(script|style|head)[^>]*>.*?</\1>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex HiddenBlocks();

    [GeneratedRegex(@"<br\s*/?>|</(tr|li)>", RegexOptions.IgnoreCase)]
    private static partial Regex LineBreaks();

    [GeneratedRegex(@"</(p|div|h[1-6]|table|ul|ol)>", RegexOptions.IgnoreCase)]
    private static partial Regex BlockEnds();

    [GeneratedRegex(@"<[^>]+>")]
    private static partial Regex Tags();

    [GeneratedRegex(@"[ \t ]+")]
    private static partial Regex SpaceRuns();

    [GeneratedRegex(@"\s*\n\s*\n\s*")]
    private static partial Regex BlankLines();
}
