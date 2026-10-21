using System.Text;

namespace Outlook.Core.Search;

/// <summary>
/// A parsed search box query, Outlook style: free words plus "from:", "to:", "subject:",
/// "category:", "has:attachment", "is:unread", "is:read", "is:flagged". Quotes group words
/// ("from:\"Ana Lopez\"", "\"exact phrase\"").
/// </summary>
public sealed record MailSearchQuery
{
    public List<string> Words { get; } = [];
    public List<string> From { get; } = [];
    public List<string> To { get; } = [];
    public List<string> Subject { get; } = [];
    public List<string> Categories { get; } = [];
    public bool? HasAttachment { get; set; }
    public bool? IsRead { get; set; }
    public bool? IsFlagged { get; set; }

    public bool IsEmpty =>
        Words.Count + From.Count + To.Count + Subject.Count + Categories.Count == 0
        && HasAttachment is null && IsRead is null && IsFlagged is null;

    public static MailSearchQuery Parse(string? text)
    {
        var query = new MailSearchQuery();
        foreach (var token in Tokenize(text ?? ""))
        {
            var colon = token.IndexOf(':');
            if (colon > 0 && colon < token.Length - 1)
            {
                var key = token[..colon].ToLowerInvariant();
                var value = token[(colon + 1)..];
                var lower = value.ToLowerInvariant();
                switch (key)
                {
                    case "from": query.From.Add(value); continue;
                    case "to": query.To.Add(value); continue;
                    case "subject": query.Subject.Add(value); continue;
                    case "category": query.Categories.Add(value); continue;
                    case "has" when lower is "attachment" or "attachments": query.HasAttachment = true; continue;
                    case "is" when lower == "unread": query.IsRead = false; continue;
                    case "is" when lower == "read": query.IsRead = true; continue;
                    case "is" when lower is "flagged" or "flag": query.IsFlagged = true; continue;
                }
            }
            query.Words.Add(token);
        }
        return query;
    }

    /// <summary>
    /// The SQLite FTS5 MATCH expression for the text parts, or null when only flags were asked for.
    /// Every term is quoted (so user input can never be read as FTS syntax) and prefix-matched,
    /// so "contr" finds "contract" the way Outlook's search-as-you-type does.
    /// </summary>
    public string? ToFtsMatch()
    {
        var parts = new List<string>();
        parts.AddRange(Words.Select(Term));
        parts.AddRange(From.Select(v => "{FromName FromAddress} : " + Term(v)));
        parts.AddRange(To.Select(v => "{To Cc} : " + Term(v)));
        parts.AddRange(Subject.Select(v => "Subject : " + Term(v)));
        return parts.Count == 0 ? null : string.Join(" AND ", parts);
    }

    private static string Term(string value) => "\"" + value.Replace("\"", "\"\"") + "\"*";

    private static IEnumerable<string> Tokenize(string text)
    {
        var current = new StringBuilder();
        var inQuotes = false;
        foreach (var c in text)
        {
            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }
            if (char.IsWhiteSpace(c) && !inQuotes)
            {
                if (current.Length > 0) yield return current.ToString();
                current.Clear();
                continue;
            }
            current.Append(c);
        }
        if (current.Length > 0) yield return current.ToString();
    }
}
