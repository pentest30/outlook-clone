using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Ical.Net.Serialization;
using Outlook.Core.Compose;
using Outlook.Core.Models;

namespace Outlook.Calendar;

/// <summary>Turning appointments into iCalendar (RFC 5545) and back.</summary>
internal static class Ics
{
    public const string ProductId = "-//OutlookClone//Calendar//EN";

    /// <summary>Outlook reads this to show Busy / Tentative / Free / Away, beyond what TRANSP can say.</summary>
    private const string BusyStatusProperty = "X-MICROSOFT-CDO-BUSYSTATUS";

    /// <summary>iCalendar wants IANA zone names ("Europe/Paris"); Windows reports its own ("Romance Standard Time").</summary>
    public static string LocalTimeZoneId()
    {
        var local = TimeZoneInfo.Local;
        if (local.HasIanaId) return local.Id;
        return TimeZoneInfo.TryConvertWindowsIdToIanaId(local.Id, out var iana) ? iana : "UTC";
    }

    public static Ical.Net.Calendar Parse(string ics) => Ical.Net.Calendar.Load(ics)
        ?? throw new FormatException("Not an iCalendar object.");

    public static string Serialize(Ical.Net.Calendar calendar) =>
        new CalendarSerializer().SerializeToString(calendar) ?? throw new InvalidOperationException("Could not write iCalendar.");

    /// <summary>The series itself (the VEVENT without RECURRENCE-ID); edited single occurrences are the others.</summary>
    public static CalendarEvent Master(Ical.Net.Calendar calendar) =>
        calendar.Events.FirstOrDefault(e => e.RecurrenceIdentifier is null) ?? calendar.Events.First();

    public static Ical.Net.Calendar Build(Account account, string uid, int sequence, AppointmentDraft draft,
        IReadOnlyDictionary<string, string>? keepStatuses = null, Ical.Net.Calendar? previous = null)
    {
        var tzId = LocalTimeZoneId();
        var calendar = new Ical.Net.Calendar { ProductId = ProductId };
        var ev = new CalendarEvent
        {
            Uid = uid,
            Sequence = sequence,
            Summary = draft.Title,
            Location = string.IsNullOrWhiteSpace(draft.Location) ? null : draft.Location,
            Description = string.IsNullOrWhiteSpace(draft.Notes) ? null : draft.Notes,
            DtStamp = CalDateTime.UtcNow,
            LastModified = CalDateTime.UtcNow,
        };

        if (draft.IsAllDay)
        {
            // All-day items are dates, not instants: the end is the day after the last day (exclusive).
            ev.DtStart = new CalDateTime(DateOnly.FromDateTime(draft.Start.Date));
            ev.DtEnd = new CalDateTime(DateOnly.FromDateTime(draft.End.Date));
        }
        else
        {
            calendar.AddTimeZone(tzId);
            ev.DtStart = new CalDateTime(draft.Start.ToLocalTime().DateTime, tzId, true);
            ev.DtEnd = new CalDateTime(draft.End.ToLocalTime().DateTime, tzId, true);
        }

        if (RuleFor(draft) is { } rule)
        {
            ev.RecurrenceRule = new RecurrenceRule(rule);
            // Keep "deleted this occurrence" exceptions when the series is edited.
            if (previous is not null)
                foreach (var exdate in Master(previous).ExceptionDates.GetAllDates())
                    ev.ExceptionDates.Add(exdate);
        }

        ev.Transparency = draft.ShowAs == ShowAs.Free ? "TRANSPARENT" : "OPAQUE";
        ev.AddProperty(BusyStatusProperty, draft.ShowAs switch
        {
            ShowAs.Free => "FREE",
            ShowAs.Tentative => "TENTATIVE",
            ShowAs.OutOfOffice => "OOF",
            _ => "BUSY",
        });

        if (draft.ReminderMinutes is { } minutes)
            ev.Alarms.Add(new Alarm { Action = "DISPLAY", Description = "Reminder", Trigger = new Trigger($"-PT{minutes}M") });

        var attendees = AddressList.Split(draft.Attendees);
        if (attendees.Count > 0)
        {
            ev.Organizer = new Organizer("mailto:" + account.Email) { CommonName = account.DisplayName };
            foreach (var entry in attendees)
            {
                var address = AddressList.AddressOf(entry);
                ev.Attendees.Add(new Attendee("mailto:" + address)
                {
                    CommonName = NameOf(entry),
                    Role = "REQ-PARTICIPANT",
                    Rsvp = true,
                    ParticipationStatus = keepStatuses?.GetValueOrDefault(address) ?? "NEEDS-ACTION",
                });
            }
        }

        calendar.Events.Add(ev);
        return calendar;
    }

    /// <summary>The RRULE for the editor's "Repeat" choice, anchored on the start date.</summary>
    public static string? RuleFor(AppointmentDraft draft)
    {
        var start = draft.Start.ToLocalTime();
        var rule = draft.Repeat switch
        {
            Repeat.Daily => "FREQ=DAILY",
            Repeat.Weekdays => "FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR",
            Repeat.Weekly => "FREQ=WEEKLY;BYDAY=" + DayCode(start.DayOfWeek),
            Repeat.Monthly => $"FREQ=MONTHLY;BYMONTHDAY={start.Day}",
            Repeat.Yearly => $"FREQ=YEARLY;BYMONTH={start.Month};BYMONTHDAY={start.Day}",
            _ => null,
        };
        if (rule is null) return null;
        if (draft.RepeatUntil is { } until)
        {
            // Inclusive of the whole last day, in UTC as RFC 5545 requires when DTSTART has a zone.
            var endOfDay = until.ToDateTime(new TimeOnly(23, 59, 59), DateTimeKind.Local).ToUniversalTime();
            rule += draft.IsAllDay ? $";UNTIL={until:yyyyMMdd}" : $";UNTIL={endOfDay:yyyyMMdd'T'HHmmss'Z'}";
        }
        return rule;
    }

    public static Repeat RepeatOf(RecurrenceRule? rule)
    {
        if (rule is null) return Repeat.Never;
        return rule.Frequency switch
        {
            FrequencyType.Daily => Repeat.Daily,
            FrequencyType.Weekly when rule.ByDay.Count == 5 => Repeat.Weekdays,
            FrequencyType.Weekly => Repeat.Weekly,
            FrequencyType.Monthly => Repeat.Monthly,
            FrequencyType.Yearly => Repeat.Yearly,
            _ => Repeat.Never,
        };
    }

    /// <summary>"Every week on Monday", for invitation summaries.</summary>
    public static string? Describe(CalendarEvent ev)
    {
        if (ev.RecurrenceRule is not { } rule) return null;
        var every = rule.Interval > 1 ? $"Every {rule.Interval} " : "Every ";
        var text = rule.Frequency switch
        {
            FrequencyType.Daily => every + (rule.Interval > 1 ? "days" : "day"),
            FrequencyType.Weekly when rule.ByDay.Count > 0 =>
                every + (rule.Interval > 1 ? "weeks" : "week") + " on " + string.Join(", ", rule.ByDay.Select(d => d.DayOfWeek)),
            FrequencyType.Weekly => every + (rule.Interval > 1 ? "weeks" : "week"),
            FrequencyType.Monthly => every + (rule.Interval > 1 ? "months" : "month"),
            FrequencyType.Yearly => every + "year",
            _ => "Repeats",
        };
        if (rule.Until is { } until) text += $" until {until.Date:d}";
        else if (rule.Count is { } count) text += $", {count} times";
        return text;
    }

    public static string DayCode(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => "MO",
        DayOfWeek.Tuesday => "TU",
        DayOfWeek.Wednesday => "WE",
        DayOfWeek.Thursday => "TH",
        DayOfWeek.Friday => "FR",
        DayOfWeek.Saturday => "SA",
        _ => "SU",
    };

    public static ShowAs ShowAsOf(CalendarEvent ev)
    {
        var status = ev.Properties.FirstOrDefault(p => p.Name == BusyStatusProperty)?.Value?.ToString();
        return status?.ToUpperInvariant() switch
        {
            "FREE" => ShowAs.Free,
            "TENTATIVE" => ShowAs.Tentative,
            "OOF" => ShowAs.OutOfOffice,
            "BUSY" => ShowAs.Busy,
            _ => ev.Transparency == "TRANSPARENT" ? ShowAs.Free : ShowAs.Busy,
        };
    }

    public static int? ReminderOf(CalendarEvent ev)
    {
        var trigger = ev.Alarms.FirstOrDefault()?.Trigger;
        if (trigger?.Duration is not { } duration) return null;
        return (int)Math.Abs(duration.ToTimeSpanUnspecified().TotalMinutes);
    }

    public static MeetingResponse ResponseOf(string? partStat) => partStat?.ToUpperInvariant() switch
    {
        "ACCEPTED" => MeetingResponse.Accepted,
        "TENTATIVE" => MeetingResponse.Tentative,
        "DECLINED" => MeetingResponse.Declined,
        "NEEDS-ACTION" => MeetingResponse.NeedsAction,
        _ => MeetingResponse.None,
    };

    public static string PartStatOf(MeetingResponse response) => response switch
    {
        MeetingResponse.Accepted => "ACCEPTED",
        MeetingResponse.Tentative => "TENTATIVE",
        MeetingResponse.Declined => "DECLINED",
        _ => "NEEDS-ACTION",
    };

    public static string EmailOf(Uri? uri) =>
        uri is null ? "" : uri.ToString().Replace("mailto:", "", StringComparison.OrdinalIgnoreCase).Trim().ToLowerInvariant();

    public static string AttendeeList(CalendarEvent ev) =>
        AddressList.Join(ev.Attendees.Select(a =>
            string.IsNullOrWhiteSpace(a.CommonName) ? EmailOf(a.Value) : $"{a.CommonName} <{EmailOf(a.Value)}>"));

    /// <summary>Start of an occurrence as an instant; all-day dates become 00:00 UTC (floating).</summary>
    public static DateTimeOffset ToInstant(CalDateTime value) =>
        value.HasTime
            ? new DateTimeOffset(DateTime.SpecifyKind(value.AsUtc, DateTimeKind.Utc))
            : new DateTimeOffset(value.Date.ToDateTime(TimeOnly.MinValue), TimeSpan.Zero);

    private static string? NameOf(string entry)
    {
        var open = entry.LastIndexOf('<');
        if (open <= 0) return null;
        var name = entry[..open].Trim().Trim('"');
        return name.Length == 0 ? null : name;
    }
}
