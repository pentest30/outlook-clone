using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Outlook.Core.Text;

/// <summary>
/// Prepares a message body for the reading pane. HtmlRenderer never runs scripts, but we still strip
/// active content, and we report remote pictures so the UI can block them (they are how senders
/// track when you open mail) until the user clicks "Download pictures".
/// </summary>
public static partial class EmailHtml
{
    public record Prepared(string Html, bool HasRemoteContent);

    public static Prepared Prepare(string? html, string? text)
    {
        if (string.IsNullOrWhiteSpace(html))
            return new Prepared(FromPlainText(text ?? ""), false);

        var clean = Sanitize(html);
        return new Prepared(clean, RemoteContent().IsMatch(clean));
    }

    public static string Sanitize(string html)
    {
        var s = DangerousBlocks().Replace(html, "");
        s = DangerousTags().Replace(s, "");
        s = EventAttributes().Replace(s, "");
        s = ScriptUrls().Replace(s, "$1=\"#\"");
        // <img src=""> is a sender bug; browsers draw nothing, but the renderer would draw a broken-image box.
        s = EmptyImages().Replace(s, "");
        return s;
    }

    /// <summary>Plain-text mail as HTML: escaped, line breaks kept, links clickable.</summary>
    public static string FromPlainText(string text)
    {
        var sb = new StringBuilder("<div style=\"font-family: Segoe UI, sans-serif; font-size: 10.5pt\">");
        foreach (var line in text.Replace("\r\n", "\n").Split('\n'))
        {
            var encoded = WebUtility.HtmlEncode(line);
            encoded = Urls().Replace(encoded, m => $"<a href=\"{m.Value}\">{m.Value}</a>");
            // Keep indentation: a leading run of spaces would otherwise collapse to one.
            var indent = encoded.Length - encoded.TrimStart(' ').Length;
            sb.Append(string.Concat(Enumerable.Repeat("&nbsp;", indent))).Append(encoded.TrimStart(' ')).Append("<br>");
        }
        return sb.Append("</div>").ToString();
    }

    [GeneratedRegex(@"<(script|iframe|object|embed|form|noscript)\b[^>]*>.*?</\1\s*>", RegexOptions.IgnoreCase | RegexOptions.Singleline)]
    private static partial Regex DangerousBlocks();

    [GeneratedRegex(@"<\/?(script|iframe|object|embed|form|base|meta|input|button|textarea|select)\b[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex DangerousTags();

    [GeneratedRegex(@"\s+on[a-z]+\s*=\s*(""[^""]*""|'[^']*'|[^\s>]+)", RegexOptions.IgnoreCase)]
    private static partial Regex EventAttributes();

    [GeneratedRegex(@"\b(href|src)\s*=\s*[""']?\s*(javascript|vbscript|data:text/html)[^""'\s>]*[""']?", RegexOptions.IgnoreCase)]
    private static partial Regex ScriptUrls();

    [GeneratedRegex(@"<img\b[^>]*\bsrc\s*=\s*(""\s*""|'\s*')[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex EmptyImages();

    // Pictures, backgrounds and stylesheets loaded from the web.
    [GeneratedRegex(@"(<img\b[^>]*\bsrc\s*=\s*[""']?\s*https?:|\bbackground\s*=\s*[""']?\s*https?:|url\(\s*[""']?\s*https?:|<link\b[^>]*\bhref\s*=\s*[""']?\s*https?:)", RegexOptions.IgnoreCase)]
    private static partial Regex RemoteContent();

    [GeneratedRegex(@"\b(https?://|mailto:)[^\s<>""']+[^\s<>""'.,;:!?)]", RegexOptions.IgnoreCase)]
    private static partial Regex Urls();

    /// <summary>True for references the reading pane must not fetch until the user allows it.</summary>
    public static bool IsRemote(string? src) =>
        src is not null && (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                            src.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                            src.StartsWith("//", StringComparison.Ordinal));
}
