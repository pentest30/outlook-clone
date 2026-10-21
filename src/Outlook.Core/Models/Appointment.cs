namespace Outlook.Core.Models;

public enum ShowAs
{
    Busy,
    Tentative,
    Free,
    OutOfOffice,
}

/// <summary>My answer to a meeting someone else organised (iCalendar PARTSTAT).</summary>
public enum MeetingResponse
{
    None,          // my own appointment, or a meeting I organise
    NeedsAction,   // invited, not answered yet
    Accepted,
    Tentative,
    Declined,
}

/// <summary>
/// A calendar item, stored as its iCalendar text (<see cref="Ics"/>, the source of truth, which
/// keeps recurrence rules, exceptions and attendees lossless) plus columns copied out of it so
/// date-range queries stay cheap. Kept locally, like classic Outlook's calendar for IMAP accounts.
/// </summary>
public class Appointment
{
    public int Id { get; set; }
    public int AccountId { get; set; }

    /// <summary>iCalendar UID: the identity shared with everyone invited.</summary>
    public string Uid { get; set; } = "";
    public int Sequence { get; set; }

    public string Title { get; set; } = "";
    public string Location { get; set; } = "";

    /// <summary>First occurrence. For all-day items the date at 00:00 UTC, read as a floating date.</summary>
    public DateTimeOffset StartUtc { get; set; }
    public DateTimeOffset EndUtc { get; set; }
    public bool IsAllDay { get; set; }

    public bool IsRecurring { get; set; }
    /// <summary>When the series stops; null for series that never end. Equals EndUtc for single items.</summary>
    public DateTimeOffset? SeriesEndUtc { get; set; }

    public string OrganizerEmail { get; set; } = "";
    public string OrganizerName { get; set; } = "";
    public bool IsOrganizer { get; set; } = true;
    /// <summary>"Ana Lopez &lt;ana@x&gt;; bob@y", for display; the iCalendar text has the details.</summary>
    public string Attendees { get; set; } = "";
    public MeetingResponse MyResponse { get; set; }

    public ShowAs ShowAs { get; set; }
    public int? ReminderMinutes { get; set; } = 15;
    /// <summary>Reminders for occurrences starting at or before this were dismissed.</summary>
    public DateTimeOffset? RemindersDismissedThroughUtc { get; set; }

    public string Ics { get; set; } = "";
    public DateTimeOffset LastModifiedUtc { get; set; }

    public bool IsMeeting => !string.IsNullOrWhiteSpace(Attendees) || !IsOrganizer;
}

/// <summary>One concrete occurrence of an appointment (the appointment itself when it does not repeat).</summary>
public sealed record AppointmentOccurrence(Appointment Appointment, DateTimeOffset Start, DateTimeOffset End)
{
    public bool IsAllDay => Appointment.IsAllDay;
}

public enum Repeat
{
    Never,
    Daily,
    Weekdays,
    Weekly,
    Monthly,
    Yearly,
}

/// <summary>What the appointment editor produces; turned into iCalendar by the calendar service.</summary>
public sealed record AppointmentDraft(
    string Title,
    string Location,
    DateTimeOffset Start,
    DateTimeOffset End,
    bool IsAllDay,
    Repeat Repeat,
    DateOnly? RepeatUntil,
    int? ReminderMinutes,
    ShowAs ShowAs,
    string Attendees,
    string Notes);
