namespace Outlook.Core.Models;

/// <summary>
/// A file part of a downloaded message, cached alongside the body. Inline parts (IsInline) are
/// pictures referenced from the HTML as cid:ContentId rather than files shown to the user.
/// </summary>
public class MailAttachment
{
    public int Id { get; set; }
    public int MessageId { get; set; }
    public MailMessage? Message { get; set; }

    public string FileName { get; set; } = "";
    public string ContentType { get; set; } = "application/octet-stream";
    public long Size { get; set; }
    public string? ContentId { get; set; }
    public bool IsInline { get; set; }

    /// <summary>Empty when loaded as metadata only; fetch the bytes with IMailStore.GetAttachmentDataAsync.</summary>
    public byte[] Data { get; set; } = [];
}
