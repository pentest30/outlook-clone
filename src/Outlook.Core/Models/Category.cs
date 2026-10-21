using System.Text;

namespace Outlook.Core.Models;

/// <summary>
/// An Outlook colour category. On the server it is an IMAP keyword on each message, so it
/// follows the mail to other devices; name and colour live in the local cache.
/// </summary>
public class Category
{
    public int Id { get; set; }
    public int AccountId { get; set; }
    public string Name { get; set; } = "";
    /// <summary>"#RRGGBB".</summary>
    public string Color { get; set; } = "#0078D4";
    public string Keyword { get; set; } = "";

    public const string KeywordPrefix = "$Cat_";

    /// <summary>Outlook's six default categories.</summary>
    public static readonly (string Name, string Color)[] Defaults =
    [
        ("Red category", "#E74856"),
        ("Orange category", "#F7630C"),
        ("Yellow category", "#FFB900"),
        ("Green category", "#16C60C"),
        ("Blue category", "#0078D4"),
        ("Purple category", "#8764B8"),
    ];

    /// <summary>
    /// IMAP keywords are atoms (no spaces, quotes or brackets), so other characters become "_".
    /// "Project: Alpha" gives "$Cat_Project__Alpha".
    /// </summary>
    public static string KeywordFor(string name)
    {
        var sb = new StringBuilder(KeywordPrefix);
        foreach (var c in name.Trim())
            sb.Append(char.IsAsciiLetterOrDigit(c) || c is '-' or '.' ? c : '_');
        return sb.ToString();
    }

    /// <summary>Best-effort display name for a keyword set by another client that we have no record of.</summary>
    public static string NameFromKeyword(string keyword) =>
        keyword.StartsWith(KeywordPrefix, StringComparison.Ordinal) ? keyword[KeywordPrefix.Length..].Replace('_', ' ') : keyword;
}
