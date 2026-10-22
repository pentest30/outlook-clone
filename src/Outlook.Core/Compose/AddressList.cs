using System.Text;

namespace Outlook.Core.Compose;

/// <summary>
/// Splits a typed recipient line ("Ana &lt;ana@x&gt;; \"Smith, Bob\" &lt;bob@y&gt;, carl@z") into single
/// entries, ignoring separators inside quotes or angle brackets. Outlook accepts both ';' and ','.
/// </summary>
public static class AddressList
{
    public static List<string> Split(string? line)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(line))
            return result;

        var current = new StringBuilder();
        var inQuotes = false;
        var inAngle = false;
        foreach (var c in line)
        {
            switch (c)
            {
                case '"': inQuotes = !inQuotes; break;
                case '<' when !inQuotes: inAngle = true; break;
                case '>' when !inQuotes: inAngle = false; break;
                case ',' or ';' when !inQuotes && !inAngle:
                    Flush();
                    continue;
            }
            current.Append(c);
        }
        Flush();
        return result;

        void Flush()
        {
            var entry = current.ToString().Trim();
            if (entry.Length > 0)
                result.Add(entry);
            current.Clear();
        }
    }

    /// <summary>The bare address of an entry: the part inside &lt;&gt;, or the whole entry.</summary>
    public static string AddressOf(string entry)
    {
        var open = entry.LastIndexOf('<');
        var close = entry.LastIndexOf('>');
        var address = open >= 0 && close > open ? entry[(open + 1)..close] : entry;
        return address.Trim().Trim('"').ToLowerInvariant();
    }

    public static string Join(IEnumerable<string> entries) => string.Join("; ", entries);

    /// <summary>"Ana Lopez &lt;ana@x&gt;" gives ("Ana Lopez", "ana@x"); a bare address has no name.</summary>
    public static (string Name, string Address) Parse(string entry)
    {
        var address = AddressOf(entry);
        var open = entry.LastIndexOf('<');
        var name = open > 0 ? entry[..open].Trim().Trim('"').Trim() : "";
        return (name, address);
    }

    /// <summary>The text being typed after the last separator, for address suggestions.</summary>
    public static string LastEntry(string? line)
    {
        if (string.IsNullOrEmpty(line)) return "";
        var cut = line.LastIndexOfAny([';', ',']);
        return line[(cut + 1)..].TrimStart();
    }

    /// <summary>The line with its last (partial) entry replaced by a completed one, ready for the next.</summary>
    public static string CompleteLast(string? line, string completed)
    {
        line ??= "";
        var cut = line.LastIndexOfAny([';', ',']);
        var head = cut < 0 ? "" : line[..(cut + 1)].TrimEnd() + " ";
        return head + completed + "; ";
    }

    /// <summary>"me@contoso.com" gives "contoso.com"; used to build Message-IDs.</summary>
    public static string DomainOf(string email)
    {
        var at = email.LastIndexOf('@');
        return at >= 0 && at < email.Length - 1 ? email[(at + 1)..] : "localhost";
    }
}
