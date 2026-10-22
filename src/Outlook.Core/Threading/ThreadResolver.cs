namespace Outlook.Core.Threading;

/// <summary>
/// Works out which conversation a message belongs to from its headers.
/// The thread root is the first Message-ID in References (the original message). If there is
/// no References header, In-Reply-To is used, and failing that the message starts its own thread.
/// </summary>
public static class ThreadResolver
{
    private static readonly string[] ReplyPrefixes = ["re", "fw", "fwd", "aw", "tr", "sv"];

    public static string Resolve(string? messageId, string? inReplyTo, string? references, string subject)
    {
        var root = FirstId(references) ?? FirstId(inReplyTo) ?? Normalize(messageId);
        if (root is not null)
            return root;

        // No usable headers at all: fall back to the subject so "Re: Lunch" joins "Lunch".
        return "subject:" + NormalizeSubject(subject);
    }

    public static string NormalizeSubject(string subject)
    {
        var s = subject.Trim();
        while (true)
        {
            var colon = s.IndexOf(':');
            if (colon is <= 0 or > 4)
                break;
            var prefix = s[..colon].Trim().ToLowerInvariant();
            if (!ReplyPrefixes.Contains(prefix))
                break;
            s = s[(colon + 1)..].TrimStart();
        }
        return s.ToLowerInvariant();
    }

    private static string? FirstId(string? header)
    {
        if (string.IsNullOrWhiteSpace(header))
            return null;
        var first = header.Split((char[])[' ', '\t', '\r', '\n', ','], StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
        return Normalize(first);
    }

    private static string? Normalize(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
            return null;
        return id.Trim().Trim('<', '>').ToLowerInvariant();
    }
}
