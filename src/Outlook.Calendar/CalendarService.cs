using System.Collections.Concurrent;
using System.Globalization;
using System.Text;
using Ical.Net;
using Ical.Net.CalendarComponents;
using Ical.Net.DataTypes;
using Microsoft.EntityFrameworkCore;
using Outlook.Core.Compose;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Data;

namespace Outlook.Calendar;

/// <summary>
/// Local calendar with iCalendar storage, plus meeting invitations over e-mail (iTIP, RFC 5546,
/// carried by iMIP, RFC 6047). See docs/PLAN.md, Phase 4.
/// </summary>
public sealed class CalendarService(Func<MailDbContext> createContext, IMailService mail) : ICalendarService
{
    /// <summary>Parsed iCalendar per appointment version, so redrawing a week does not re-parse every series.</summary>
    private readonly ConcurrentDictionary<(int Id, long Version), Ical.Net.Calendar> _parsed = new();

    // ───────────────────────────── Reading ─────────────────────────────

    public async Task<IReadOnlyList<AppointmentOccurrence>> GetOccurrencesAsync(int accountId, DateTimeOffset from, DateTimeOffset to,
        CancellationToken ct = default)
    {
        await using var db = createContext();
        // Single items have SeriesEndUtc == EndUtc, endless series null, so one range query covers both.
        var candidates = await db.Appointments.AsNoTracking()
            .Where(a => a.AccountId == accountId && a.StartUtc < to && (a.SeriesEndUtc == null || a.SeriesEndUtc > from))
            .ToListAsync(ct);

        var result = new List<AppointmentOccurrence>();
        foreach (var a in candidates)
        {
            if (!a.IsRecurring)
            {
                result.Add(new AppointmentOccurrence(a, a.StartUtc, a.EndUtc));
                continue;
            }
            result.AddRange(Expand(a, from, to));
        }
        return result.Where(o => o.Start < to && o.End > from).OrderBy(o => o.Start).ThenBy(o => o.End).ToList();
    }

    private IEnumerable<AppointmentOccurrence> Expand(Appointment a, DateTimeOffset from, DateTimeOffset to)
    {
        var calendar = _parsed.GetOrAdd((a.Id, a.LastModifiedUtc.UtcTicks), _ => Ics.Parse(a.Ics));
        var length = a.EndUtc - a.StartUtc;
        // Start early enough to catch an occurrence that began before the range but is still running.
        var windowStart = new CalDateTime((from - length - TimeSpan.FromDays(1)).UtcDateTime, "UTC", true);
        var windowEnd = new CalDateTime(to.UtcDateTime, "UTC", true);

        foreach (var occurrence in calendar.GetOccurrences(windowStart, null).TakeWhileBefore(windowEnd))
        {
            var start = Ics.ToInstant(occurrence.Period.StartTime);
            var end = occurrence.Period.EffectiveEndTime is { } e ? Ics.ToInstant(e) : start + length;
            yield return new AppointmentOccurrence(a, start, end);
        }
    }

    public async Task<Appointment?> GetAsync(int appointmentId, CancellationToken ct = default)
    {
        await using var db = createContext();
        return await db.Appointments.AsNoTracking().SingleOrDefaultAsync(a => a.Id == appointmentId, ct);
    }

    public async Task<Appointment?> FindAsync(int accountId, string uid, CancellationToken ct = default)
    {
        await using var db = createContext();
        return await db.Appointments.AsNoTracking().SingleOrDefaultAsync(a => a.AccountId == accountId && a.Uid == uid, ct);
    }

    public AppointmentDraft ToDraft(Appointment a)
    {
        var ev = Ics.Master(Ics.Parse(a.Ics));
        var rule = ev.RecurrenceRule;
        DateOnly? until = rule?.Until is { } u ? DateOnly.FromDateTime(a.IsAllDay ? u.Value : u.AsUtc.ToLocalTime()) : null;
        return new AppointmentDraft(a.Title, a.Location, a.StartUtc, a.EndUtc, a.IsAllDay, Ics.RepeatOf(rule), until,
            a.ReminderMinutes, a.ShowAs, a.Attendees, ev.Description ?? "");
    }

    // ───────────────────────────── My appointments ─────────────────────────────

    public async Task<(Appointment Saved, string? Problem)> SaveAsync(Account account, Appointment? existing, AppointmentDraft draft,
        CancellationToken ct = default)
    {
        if (draft.End <= draft.Start)
            throw new ArgumentException("The end must be after the start.", nameof(draft));

        await using var db = createContext();
        var row = existing is null ? null : await db.Appointments.SingleOrDefaultAsync(a => a.Id == existing.Id, ct);
        var uid = row?.Uid ?? $"{Guid.NewGuid():N}@{AddressList.DomainOf(account.Email)}";
        var sequence = row is null ? 0 : row.Sequence + 1;   // RFC 5546: bump on every change attendees should see

        Ical.Net.Calendar? previous = row is null ? null : Ics.Parse(row.Ics);
        // Keep the answers attendees already sent; they only reset if the meeting is rebuilt from scratch.
        var statuses = previous is null
            ? null
            : Ics.Master(previous).Attendees.ToDictionary(x => Ics.EmailOf(x.Value), x => x.ParticipationStatus ?? "NEEDS-ACTION");
        var calendar = Ics.Build(account, uid, sequence, draft, statuses, previous);

        row ??= db.Appointments.Add(new Appointment { AccountId = account.Id, Uid = uid }).Entity;
        Fill(row, calendar, account.Email);
        row.IsOrganizer = true;
        row.MyResponse = MeetingResponse.None;
        await db.SaveChangesAsync(ct);

        var problem = row.IsMeeting
            ? await SendAsync(account, calendar, "REQUEST", AddressList.Split(draft.Attendees), row.Title,
                existing is null ? "invitation" : "update", ct)
            : null;
        return (row, problem);
    }

    public async Task<string?> DeleteAsync(Account account, Appointment appointment, DateTimeOffset? occurrenceStart, CancellationToken ct = default)
    {
        await using var db = createContext();
        var row = await db.Appointments.SingleOrDefaultAsync(a => a.Id == appointment.Id, ct);
        if (row is null) return null;
        var calendar = Ics.Parse(row.Ics);
        var master = Ics.Master(calendar);
        var notify = row.IsOrganizer && master.Attendees.Count > 0;

        if (occurrenceStart is { } start && row.IsRecurring)
        {
            // One occurrence: an EXDATE in the series (and the cancellation names just that occurrence).
            var exdate = row.IsAllDay
                ? new CalDateTime(DateOnly.FromDateTime(start.UtcDateTime))
                : new CalDateTime(start.UtcDateTime, "UTC", true).ToTimeZone(master.DtStart!.TzId);
            master.ExceptionDates.Add(exdate);
            master.Sequence = ++row.Sequence;
            row.Ics = Ics.Serialize(calendar);
            row.LastModifiedUtc = DateTimeOffset.UtcNow;
            await db.SaveChangesAsync(ct);
            return notify ? await SendCancelAsync(account, row, master, exdate, ct) : null;
        }

        db.Appointments.Remove(row);
        await db.SaveChangesAsync(ct);
        return notify ? await SendCancelAsync(account, row, master, null, ct) : null;
    }

    private Task<string?> SendCancelAsync(Account account, Appointment row, CalendarEvent master, CalDateTime? occurrence, CancellationToken ct)
    {
        var cancel = new Ical.Net.Calendar { ProductId = Ics.ProductId };
        var ev = master.Copy<CalendarEvent>() ?? throw new InvalidOperationException("Could not copy the event.");
        ev.Sequence = row.Sequence + 1;
        ev.Status = "CANCELLED";
        ev.DtStamp = CalDateTime.UtcNow;
        if (occurrence is not null)
        {
            ev.RecurrenceRule = null;
            ev.ExceptionDates.Clear();
            ev.RecurrenceIdentifier = new RecurrenceIdentifier(occurrence);
            var length = ev.DtEnd!.Subtract(ev.DtStart!);
            ev.DtStart = occurrence;
            ev.DtEnd = occurrence.Add(length);
        }
        cancel.Events.Add(ev);
        var recipients = master.Attendees.Select(a => Ics.EmailOf(a.Value)).ToList();
        return SendAsync(account, cancel, "CANCEL", recipients, "Canceled: " + row.Title, "cancellation", ct);
    }

    public async Task DismissReminderAsync(Appointment appointment, DateTimeOffset occurrenceStart, CancellationToken ct = default)
    {
        await using var db = createContext();
        var row = await db.Appointments.SingleOrDefaultAsync(a => a.Id == appointment.Id, ct);
        if (row is null) return;
        if (row.RemindersDismissedThroughUtc is null || row.RemindersDismissedThroughUtc < occurrenceStart)
            row.RemindersDismissedThroughUtc = occurrenceStart;
        await db.SaveChangesAsync(ct);
    }

    // ───────────────────────────── Invitations ─────────────────────────────

    public Invitation? ParseInvitation(string ics)
    {
        Ical.Net.Calendar calendar;
        try
        {
            calendar = Ics.Parse(ics);
        }
        catch (Exception)
        {
            return null;   // broken attachments happen; the mail still reads fine without the card
        }
        if (calendar.Events.Count == 0) return null;

        var ev = Ics.Master(calendar);
        var method = string.IsNullOrWhiteSpace(calendar.Method) ? "REQUEST" : calendar.Method.ToUpperInvariant();
        var start = Ics.ToInstant(ev.DtStart!);
        var end = ev.DtEnd is { } dtEnd ? Ics.ToInstant(dtEnd) : start + (ev.Duration?.ToTimeSpan(ev.DtStart!) ?? TimeSpan.FromHours(1));
        var replier = method == "REPLY" ? ev.Attendees.FirstOrDefault() : null;

        return new Invitation(method, ev.Uid ?? "", ev.Sequence, ev.Summary ?? "(no subject)", ev.Location ?? "",
            start, end, ev.IsAllDay, Ics.Describe(ev),
            ev.Organizer?.CommonName ?? "", Ics.EmailOf(ev.Organizer?.Value), Ics.AttendeeList(ev),
            replier is null ? null : replier.CommonName ?? Ics.EmailOf(replier.Value),
            Ics.ResponseOf(replier?.ParticipationStatus), ics);
    }

    public async Task<string?> RespondAsync(Account account, Invitation invitation, MeetingResponse response, bool sendReply,
        CancellationToken ct = default)
    {
        var me = account.Email.ToLowerInvariant();
        var calendar = Ics.Parse(invitation.Ics);
        calendar.Method = null;
        foreach (var ev in calendar.Events)
        {
            var attendee = ev.Attendees.FirstOrDefault(a => Ics.EmailOf(a.Value) == me);
            if (attendee is null)
            {
                // Invited through a list or an alias: answer as myself anyway.
                attendee = new Attendee("mailto:" + me) { CommonName = account.DisplayName };
                ev.Attendees.Add(attendee);
            }
            attendee.ParticipationStatus = Ics.PartStatOf(response);
            attendee.Rsvp = false;
        }

        await using (var db = createContext())
        {
            var row = await db.Appointments.SingleOrDefaultAsync(a => a.AccountId == account.Id && a.Uid == invitation.Uid, ct);
            if (response == MeetingResponse.Declined)
            {
                // Like Outlook: a declined meeting leaves the calendar.
                if (row is not null) db.Appointments.Remove(row);
            }
            else
            {
                row ??= db.Appointments.Add(new Appointment { AccountId = account.Id, Uid = invitation.Uid }).Entity;
                Fill(row, calendar, me);
                row.IsOrganizer = false;
                row.MyResponse = response;
                if (response == MeetingResponse.Tentative) row.ShowAs = ShowAs.Tentative;
                row.ReminderMinutes ??= 15;
            }
            await db.SaveChangesAsync(ct);
        }

        if (!sendReply || string.IsNullOrWhiteSpace(invitation.OrganizerEmail))
            return null;

        // The REPLY carries just me as attendee, with my answer (RFC 5546 section 3.2.3).
        var reply = new Ical.Net.Calendar { ProductId = Ics.ProductId };
        var master = Ics.Master(calendar);
        var answer = new CalendarEvent
        {
            Uid = invitation.Uid,
            Sequence = master.Sequence,
            Summary = master.Summary,
            DtStart = master.DtStart,
            DtEnd = master.DtEnd,
            DtStamp = CalDateTime.UtcNow,
            Organizer = master.Organizer,
        };
        answer.Attendees.Add(new Attendee("mailto:" + me) { CommonName = account.DisplayName, ParticipationStatus = Ics.PartStatOf(response) });
        reply.Events.Add(answer);

        var verb = response switch
        {
            MeetingResponse.Accepted => "Accepted",
            MeetingResponse.Tentative => "Tentative",
            _ => "Declined",
        };
        return await SendAsync(account, reply, "REPLY", [invitation.OrganizerEmail], $"{verb}: {invitation.Title}", "reply", ct);
    }

    public async Task ApplyAsync(Account account, Invitation invitation, CancellationToken ct = default)
    {
        await using var db = createContext();
        var row = await db.Appointments.SingleOrDefaultAsync(a => a.AccountId == account.Id && a.Uid == invitation.Uid, ct);
        if (row is null) return;

        if (invitation.Method == "CANCEL")
        {
            if (invitation.Sequence < row.Sequence) return;   // an older cancellation than what we have
            var cancelled = Ics.Parse(invitation.Ics);
            var ev = Ics.Master(cancelled);
            if (ev.RecurrenceIdentifier?.StartTime is { } occurrence && row.IsRecurring)
            {
                var calendar = Ics.Parse(row.Ics);
                Ics.Master(calendar).ExceptionDates.Add(occurrence);
                row.Ics = Ics.Serialize(calendar);
                row.LastModifiedUtc = DateTimeOffset.UtcNow;
            }
            else
            {
                db.Appointments.Remove(row);
            }
        }
        else if (invitation.Method == "REPLY" && row.IsOrganizer)
        {
            // Record the attendee's answer on my copy of the meeting.
            var reply = Ics.Master(Ics.Parse(invitation.Ics)).Attendees.FirstOrDefault();
            if (reply is null) return;
            var calendar = Ics.Parse(row.Ics);
            foreach (var attendee in Ics.Master(calendar).Attendees.Where(a => Ics.EmailOf(a.Value) == Ics.EmailOf(reply.Value)))
                attendee.ParticipationStatus = reply.ParticipationStatus;
            row.Ics = Ics.Serialize(calendar);
            row.LastModifiedUtc = DateTimeOffset.UtcNow;
        }
        await db.SaveChangesAsync(ct);
    }

    /// <summary>Attendee answers recorded on a meeting I organise (from their REPLY e-mails).</summary>
    public static IReadOnlyList<(string Who, MeetingResponse Response)> AttendeeResponses(Appointment appointment) =>
        Ics.Master(Ics.Parse(appointment.Ics)).Attendees
            .Select(a => (string.IsNullOrWhiteSpace(a.CommonName) ? Ics.EmailOf(a.Value) : a.CommonName!, Ics.ResponseOf(a.ParticipationStatus)))
            .ToList();

    // ───────────────────────────── Plumbing ─────────────────────────────

    /// <summary>Copies the searchable columns out of the iCalendar text.</summary>
    private static void Fill(Appointment row, Ical.Net.Calendar calendar, string myEmail)
    {
        var ev = Ics.Master(calendar);
        row.Uid = ev.Uid ?? row.Uid;
        row.Sequence = ev.Sequence;
        row.Title = ev.Summary ?? "(no subject)";
        row.Location = ev.Location ?? "";
        row.IsAllDay = ev.IsAllDay;
        row.StartUtc = Ics.ToInstant(ev.DtStart!);
        row.EndUtc = ev.DtEnd is { } end ? Ics.ToInstant(end) : row.StartUtc + (ev.Duration?.ToTimeSpan(ev.DtStart!) ?? TimeSpan.FromHours(1));
        row.IsRecurring = ev.RecurrenceRule is not null || ev.RecurrenceDates.GetAllDates().Any();
        row.SeriesEndUtc = row.IsRecurring ? SeriesEnd(calendar, row) : row.EndUtc;
        row.OrganizerEmail = Ics.EmailOf(ev.Organizer?.Value);
        row.OrganizerName = ev.Organizer?.CommonName ?? "";
        row.Attendees = Ics.AttendeeList(ev);
        row.ShowAs = Ics.ShowAsOf(ev);
        row.ReminderMinutes = Ics.ReminderOf(ev);
        row.Ics = Ics.Serialize(calendar);
        row.LastModifiedUtc = DateTimeOffset.UtcNow;
    }

    /// <summary>End of the last occurrence, or null for a series that never ends.</summary>
    private static DateTimeOffset? SeriesEnd(Ical.Net.Calendar calendar, Appointment row)
    {
        var rule = Ics.Master(calendar).RecurrenceRule;
        if (rule is not null && rule.Until is null && rule.Count is null)
            return null;
        var length = row.EndUtc - row.StartUtc;
        var last = calendar.GetOccurrences(null, null).Take(10_000).LastOrDefault();
        return last is null ? row.EndUtc : Ics.ToInstant(last.Period.StartTime) + length;
    }

    private async Task<string?> SendAsync(Account account, Ical.Net.Calendar calendar, string method, IReadOnlyList<string> to,
        string subject, string what, CancellationToken ct)
    {
        if (to.Count == 0) return null;
        calendar.Method = method;
        var ev = Ics.Master(calendar);
        var text = new StringBuilder();
        text.AppendLine(ev.Summary);
        text.AppendLine();
        var start = Ics.ToInstant(ev.DtStart!);
        text.AppendLine(ev.IsAllDay
            ? $"When: {start.UtcDateTime.ToString("D", CultureInfo.CurrentCulture)} (all day)"
            : $"When: {start.ToLocalTime().ToString("f", CultureInfo.CurrentCulture)} ({TimeZoneInfo.Local.DisplayName})");
        if (!string.IsNullOrWhiteSpace(ev.Location)) text.AppendLine($"Where: {ev.Location}");
        if (Ics.Describe(ev) is { } repeat) text.AppendLine($"Repeats: {repeat}");
        if (!string.IsNullOrWhiteSpace(ev.Description)) text.AppendLine().AppendLine(ev.Description);

        try
        {
            var report = await mail.SendAsync(account,
                new OutgoingMessage(to, [], [], subject, text.ToString(), Calendar: new OutgoingCalendar(Ics.Serialize(calendar), method)), ct);
            return report.Problem;
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return $"Saved, but the {what} e-mail couldn't be sent: {ex.Message}";
        }
    }
}
