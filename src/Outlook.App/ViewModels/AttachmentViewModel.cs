using Outlook.Core.Models;

namespace Outlook.App.ViewModels;

/// <summary>An attachment chip, in the reading pane (cached attachment) or the compose window (file to send).</summary>
public class AttachmentViewModel
{
    public AttachmentViewModel(MailAttachment attachment)
    {
        Attachment = attachment;
        FileName = attachment.FileName;
        Size = attachment.Size;
    }

    public AttachmentViewModel(OutgoingAttachment outgoing)
    {
        Outgoing = outgoing;
        FileName = outgoing.FileName;
        Size = outgoing.Data.LongLength;
    }

    public MailAttachment? Attachment { get; }
    public OutgoingAttachment? Outgoing { get; }
    public string FileName { get; }
    public long Size { get; }

    public string Icon => Path.GetExtension(FileName).ToLowerInvariant() switch
    {
        ".pdf" => "📕",
        ".doc" or ".docx" or ".txt" or ".rtf" => "📄",
        ".xls" or ".xlsx" or ".csv" => "📊",
        ".ppt" or ".pptx" => "📙",
        ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => "🖼",
        ".zip" or ".7z" or ".rar" => "🗜",
        ".eml" or ".msg" => "✉",
        _ => "📎",
    };

    public string SizeText => FormatSize(Size);

    public static string FormatSize(long bytes) => bytes switch
    {
        < 1024 => $"{bytes} B",
        < 1024 * 1024 => $"{bytes / 1024.0:0.#} KB",
        _ => $"{bytes / (1024.0 * 1024):0.#} MB",
    };
}
