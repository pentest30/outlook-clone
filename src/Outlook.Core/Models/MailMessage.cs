namespace Outlook.Core.Models;

public class MailMessage
{
    public int Id { get; set; }
    public int FolderId { get; set; }
    public MailFolder? Folder { get; set; }

    /// <summary>IMAP UID, only meaningful together with the folder's UidValidity.</summary>
    public uint Uid { get; set; }

    public string? MessageIdHeader { get; set; }
    public string? InReplyTo { get; set; }
    /// <summary>Space-separated Message-IDs from the References header.</summary>
    public string? References { get; set; }
    public string ThreadId { get; set; } = "";

    public string Subject { get; set; } = "";
    public string FromName { get; set; } = "";
    public string FromAddress { get; set; } = "";
    /// <summary>Formatted recipient lists, e.g. "Ana &lt;ana@x.com&gt;, bob@y.com".</summary>
    public string To { get; set; } = "";
    public string Cc { get; set; } = "";
    /// <summary>Reply-To header, when the sender wants answers to go somewhere other than From.</summary>
    public string ReplyTo { get; set; } = "";

    public DateTimeOffset Date { get; set; }
    public long Size { get; set; }
    public string Preview { get; set; } = "";

    public bool IsRead { get; set; }
    public bool IsFlagged { get; set; }
    public bool IsAnswered { get; set; }
    public bool IsForwarded { get; set; }
    public bool IsDraft { get; set; }
    public bool HasAttachments { get; set; }
    /// <summary>Carries a meeting invitation, update or reply (a text/calendar part), known from the headers.</summary>
    public bool HasCalendar { get; set; }

    /// <summary>Space-separated category keywords (see <see cref="Category.Keyword"/>).</summary>
    public string CategoryKeywords { get; set; } = "";

    /// <summary>Bodies download on first open; null until then.</summary>
    public string? BodyText { get; set; }
    public string? BodyHtml { get; set; }
    public bool IsBodyDownloaded { get; set; }
    /// <summary>Plain text of the body (from the HTML when there is no text part), for the search index.</summary>
    public string? SearchText { get; set; }
    /// <summary>The iCalendar text of that part, once the body is downloaded.</summary>
    public string? CalendarIcs { get; set; }

    public List<MailAttachment> Attachments { get; set; } = [];

    public string SenderDisplay => string.IsNullOrWhiteSpace(FromName) ? FromAddress : FromName;

    public IEnumerable<string> CategoryKeywordList => CategoryKeywords.Split(' ', StringSplitOptions.RemoveEmptyEntries);
}
