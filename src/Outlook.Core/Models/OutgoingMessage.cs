namespace Outlook.Core.Models;

/// <summary>
/// A message being sent or saved as a draft. Each entry in To/Cc/Bcc may hold one address or a
/// comma/semicolon separated list, as typed in the compose window.
/// </summary>
/// <param name="MessageId">
/// Stable for the whole compose session, so each draft save replaces the previous copy in Drafts
/// and sending can remove it. Generated when omitted.
/// </param>
public record OutgoingMessage(
    IReadOnlyList<string> To,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Bcc,
    string Subject,
    string BodyText,
    string? BodyHtml = null,
    string? InReplyTo = null,
    string? References = null,
    IReadOnlyList<OutgoingAttachment>? Attachments = null,
    string? MessageId = null,
    OutgoingCalendar? Calendar = null);

public record OutgoingAttachment(string FileName, string ContentType, byte[] Data);

/// <summary>A meeting invitation, update, cancellation or reply carried as text/calendar (iMIP, RFC 6047).</summary>
public record OutgoingCalendar(string Ics, string Method);
