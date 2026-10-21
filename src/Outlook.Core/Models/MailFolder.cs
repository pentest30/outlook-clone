namespace Outlook.Core.Models;

/// <summary>Well-known folders, mapped from IMAP SPECIAL-USE flags (RFC 6154) or common names.</summary>
public enum SpecialFolder
{
    None,
    Inbox,
    Drafts,
    Sent,
    Archive,
    Junk,
    Trash,
}

public class MailFolder
{
    public int Id { get; set; }
    public int AccountId { get; set; }
    public Account? Account { get; set; }

    /// <summary>Full IMAP path, e.g. "INBOX/Projects". Unique per account.</summary>
    public string FullName { get; set; } = "";
    public string DisplayName { get; set; } = "";
    public string? ParentFullName { get; set; }
    public SpecialFolder Special { get; set; }

    /// <summary>When the server changes this, every cached UID in the folder is invalid.</summary>
    public uint UidValidity { get; set; }
    public int UnreadCount { get; set; }
    public int TotalCount { get; set; }
    public DateTimeOffset? LastSyncedAt { get; set; }

    public List<MailMessage> Messages { get; set; } = [];

    /// <summary>Folder pane order: special folders first in Outlook's order, then alphabetical.</summary>
    public int SortRank => Special switch
    {
        SpecialFolder.Inbox => 0,
        SpecialFolder.Drafts => 1,
        SpecialFolder.Sent => 2,
        SpecialFolder.Trash => 3,
        SpecialFolder.Junk => 4,
        SpecialFolder.Archive => 5,
        _ => 10,
    };
}
