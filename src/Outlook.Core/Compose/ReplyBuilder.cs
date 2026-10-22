using System.Globalization;
using System.Text;
using Outlook.Core.Models;

namespace Outlook.Core.Compose;

public enum ComposeMode
{
    New,
    Reply,
    ReplyAll,
    Forward,
    EditDraft,
}

/// <summary>Initial contents of a compose window.</summary>
public record ComposeSeed(
    string To,
    string Cc,
    string Subject,
    string Body,
    string? InReplyTo,
    string? References);

/// <summary>
/// Pre-fills reply, reply-all and forward the way classic Outlook does: "RE:" / "FW:" subjects,
/// the original under a "From / Sent / To / Subject" header block, and threading headers so the
/// answer lands in the same conversation everywhere.
/// </summary>
public static class ReplyBuilder
{
    public const string Separator = "________________________________";

    /// <param name="signature">Added under the cursor position: alone in new mail, above the quoted original in replies.</param>
    public static ComposeSeed Create(ComposeMode mode, MailMessage? original, string originalBody, string myAddress, string signature = "")
    {
        var signed = string.IsNullOrWhiteSpace(signature) ? "" : "\n\n" + signature.TrimEnd();
        if (original is null || mode == ComposeMode.New)
            return new ComposeSeed("", "", "", signed, null, null);

        if (mode == ComposeMode.EditDraft)
            return new ComposeSeed(original.To, original.Cc, original.Subject, originalBody, original.InReplyTo, original.References);

        var me = myAddress.ToLowerInvariant();
        var references = JoinReferences(original.References, original.MessageIdHeader);
        var quoted = signed + "\n\n" + QuoteHeader(original) + originalBody;

        switch (mode)
        {
            case ComposeMode.Forward:
                return new ComposeSeed("", "", Prefix("FW:", original.Subject), quoted, null, references);

            case ComposeMode.Reply:
                return new ComposeSeed(ReplyTarget(original), "", Prefix("RE:", original.Subject), quoted,
                    original.MessageIdHeader, references);

            default: // ReplyAll
                var to = Distinct([ReplyTarget(original), .. AddressList.Split(original.To)], exclude: me);
                var cc = Distinct(AddressList.Split(original.Cc), exclude: me, alsoExclude: to);
                return new ComposeSeed(AddressList.Join(to), AddressList.Join(cc), Prefix("RE:", original.Subject),
                    quoted, original.MessageIdHeader, references);
        }
    }

    /// <summary>Adds "RE:" or "FW:" unless the subject already starts with it.</summary>
    public static string Prefix(string prefix, string subject)
    {
        var s = subject.Trim();
        return s.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? s : $"{prefix} {s}";
    }

    private static string ReplyTarget(MailMessage m)
    {
        if (!string.IsNullOrWhiteSpace(m.ReplyTo))
            return m.ReplyTo;
        return string.IsNullOrWhiteSpace(m.FromName) ? m.FromAddress : $"{m.FromName} <{m.FromAddress}>";
    }

    private static string QuoteHeader(MailMessage m)
    {
        var sb = new StringBuilder();
        sb.AppendLine(Separator);
        sb.AppendLine($"From: {ReplyTargetDisplay(m)}");
        sb.AppendLine($"Sent: {m.Date.ToLocalTime().ToString("f", CultureInfo.CurrentCulture)}");
        sb.AppendLine($"To: {m.To}");
        if (!string.IsNullOrWhiteSpace(m.Cc))
            sb.AppendLine($"Cc: {m.Cc}");
        sb.AppendLine($"Subject: {m.Subject}");
        sb.AppendLine();
        return sb.ToString();
    }

    private static string ReplyTargetDisplay(MailMessage m) =>
        string.IsNullOrWhiteSpace(m.FromName) ? m.FromAddress : $"{m.FromName} <{m.FromAddress}>";

    private static List<string> Distinct(IEnumerable<string> entries, string exclude, IEnumerable<string>? alsoExclude = null)
    {
        var seen = new HashSet<string>(alsoExclude?.Select(AddressList.AddressOf) ?? []) { exclude };
        var result = new List<string>();
        foreach (var e in entries.Where(e => !string.IsNullOrWhiteSpace(e)))
            if (seen.Add(AddressList.AddressOf(e)))
                result.Add(e);
        return result;
    }

    private static string? JoinReferences(string? references, string? messageId)
    {
        var ids = (references ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries).ToList();
        if (!string.IsNullOrWhiteSpace(messageId) && !ids.Contains(messageId))
            ids.Add(messageId);
        return ids.Count == 0 ? null : string.Join(' ', ids);
    }
}
