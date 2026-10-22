using Outlook.Core.Models;

namespace Outlook.Core.Services;

/// <summary>A meeting invitation, update, cancellation or reply found in an e-mail.</summary>
public sealed record Invitation(
    string Method,               // REQUEST, CANCEL or REPLY (iTIP, RFC 5546)
    string Uid,
    int Sequence,
    string Title,
    string Location,
    DateTimeOffset Start,
    DateTimeOffset End,
    bool IsAllDay,
    string? RecurrenceText,      // "Every week on Monday", null when it does not repeat
    string OrganizerName,
    string OrganizerEmail,
    string Attendees,
    string? ReplyFrom,           // for REPLY: who answered
    MeetingResponse ReplyStatus, // for REPLY: what they answered
    string Ics);

/// <summary>
/// The calendar: local storage with recurrence, plus meeting invitations sent and received by
/// e-mail (iTIP over iMIP), which works with any mail account.
/// </summary>
public interface ICalendarService
{
    /// <summary>Every occurrence overlapping [from, to), repeating series expanded, ordered by start.</summary>
    Task<IReadOnlyList<AppointmentOccurrence>> GetOccurrencesAsync(int accountId, DateTimeOffset from, DateTimeOffset to, CancellationToken ct = default);

    Task<Appointment?> GetAsync(int appointmentId, CancellationToken ct = default);

    /// <summary>The editor's fields for an existing appointment.</summary>
    AppointmentDraft ToDraft(Appointment appointment);

    /// <summary>
    /// Creates (existing null) or updates an appointment I organise. With attendees, sends them the
    /// invitation or update. Returns the saved appointment and any problem sending invitations.
    /// </summary>
    Task<(Appointment Saved, string? Problem)> SaveAsync(Account account, Appointment? existing, AppointmentDraft draft, CancellationToken ct = default);

    /// <summary>
    /// Deletes one occurrence of a series (pass its <see cref="AppointmentOccurrence.RecurrenceId"/>) or the
    /// whole item (null); tells attendees if I organise it.
    /// </summary>
    Task<string?> DeleteAsync(Account account, Appointment appointment, DateTimeOffset? occurrenceId, CancellationToken ct = default);

    /// <summary>
    /// Gives an occurrence a new start and end (drag to move or resize). For a series only that
    /// occurrence changes, as in Outlook. Meetings I organise send attendees an update; meetings
    /// organised by someone else cannot be moved.
    /// </summary>
    Task<string?> MoveAsync(Account account, AppointmentOccurrence occurrence, DateTimeOffset newStart, DateTimeOffset newEnd, CancellationToken ct = default);

    /// <summary>Reads the calendar part of a downloaded message, or null if it has none.</summary>
    Invitation? ParseInvitation(string ics);

    /// <summary>The appointment an invitation refers to, if it is already in the calendar.</summary>
    Task<Appointment?> FindAsync(int accountId, string uid, CancellationToken ct = default);

    /// <summary>Accept / tentative / decline: updates the calendar and (optionally) e-mails the organiser.</summary>
    Task<string?> RespondAsync(Account account, Invitation invitation, MeetingResponse response, bool sendReply, CancellationToken ct = default);

    /// <summary>Applies a CANCEL (removes the meeting) or a REPLY (records the attendee's answer).</summary>
    Task ApplyAsync(Account account, Invitation invitation, CancellationToken ct = default);

    Task DismissReminderAsync(Appointment appointment, DateTimeOffset occurrenceStart, CancellationToken ct = default);
}
