using Outlook.Calendar;
using Outlook.Core.Models;
using Outlook.Core.Services;
using Outlook.Data;

namespace Outlook.Tests;

/// <summary>Captures outgoing mail instead of sending it.</summary>
internal sealed class RecordingMailService : IMailService
{
    public List<OutgoingMessage> Sent { get; } = [];

    public Task<SendReport> SendAsync(Account account, OutgoingMessage message, CancellationToken ct = default)
    {
        Sent.Add(message);
        return Task.FromResult(new SendReport(true, null));
    }

    public event EventHandler<FolderChangedEventArgs>? FolderChanged { add { } remove { } }
    public Task VerifyAsync(Account a, string p, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SyncAccountAsync(Account a, IProgress<SyncProgress>? p = null, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SyncFolderAsync(Account a, MailFolder f, CancellationToken ct = default) => throw new NotSupportedException();
    public Task RunPushAsync(Account a, CancellationToken ct) => throw new NotSupportedException();
    public Task<MailMessage> DownloadBodyAsync(MailMessage m, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SetReadAsync(MailMessage m, bool r, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SetFlaggedAsync(MailMessage m, bool f, CancellationToken ct = default) => throw new NotSupportedException();
    public Task MarkRespondedAsync(MailMessage m, bool f, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SetCategoryAsync(IReadOnlyList<MailMessage> m, string k, bool a, CancellationToken ct = default) => throw new NotSupportedException();
    public Task MoveAsync(IReadOnlyList<MailMessage> m, MailFolder t, CancellationToken ct = default) => throw new NotSupportedException();
    public Task DeleteAsync(IReadOnlyList<MailMessage> m, CancellationToken ct = default) => throw new NotSupportedException();
    public Task ArchiveAsync(IReadOnlyList<MailMessage> m, CancellationToken ct = default) => throw new NotSupportedException();
    public Task SaveDraftAsync(Account a, OutgoingMessage d, CancellationToken ct = default) => throw new NotSupportedException();
    public Task DeleteDraftAsync(Account a, string id, CancellationToken ct = default) => throw new NotSupportedException();
}

public sealed class CalendarTests : IDisposable
{
    private readonly string _dbPath = Path.Combine(Path.GetTempPath(), $"outlook-cal-{Guid.NewGuid():N}.db");
    private readonly MailStore _store;
    private readonly RecordingMailService _mail = new();
    private readonly CalendarService _calendar;
    private readonly Account _me;
    private readonly Account _ana;

    public CalendarTests()
    {
        _store = new MailStore(() => MailDbContext.Open(_dbPath));
        _store.Migrate();
        _calendar = new CalendarService(_store.CreateContext, _mail);
        _me = _store.AddAccountAsync(new Account { Email = "me@x.test", DisplayName = "Me" }).GetAwaiter().GetResult();
        _ana = _store.AddAccountAsync(new Account { Email = "ana@x.test", DisplayName = "Ana" }).GetAwaiter().GetResult();
    }

    private static DateTimeOffset Local(int y, int mo, int d, int h, int mi = 0) =>
        new(new DateTime(y, mo, d, h, mi, 0, DateTimeKind.Local));

    private static AppointmentDraft Draft(string title, DateTimeOffset start, TimeSpan length, Repeat repeat = Repeat.Never,
        DateOnly? until = null, string attendees = "", bool allDay = false) =>
        new(title, "Room 1", start, start + length, allDay, repeat, until, 15, ShowAs.Busy, attendees, "Notes");

    [Fact]
    public async Task Single_appointment_round_trips_and_shows_in_range()
    {
        var start = Local(2026, 11, 2, 9);
        var (saved, problem) = await _calendar.SaveAsync(_me, null, Draft("Standup", start, TimeSpan.FromMinutes(30)));

        Assert.Null(problem);
        Assert.Empty(_mail.Sent);   // no attendees, no e-mail
        var day = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 3, 0));
        var occurrence = Assert.Single(day);
        Assert.Equal(start, occurrence.Start);
        Assert.Equal(start.AddMinutes(30), occurrence.End);

        var draft = _calendar.ToDraft(saved);
        Assert.Equal("Standup", draft.Title);
        Assert.Equal("Notes", draft.Notes);
        Assert.Equal(15, draft.ReminderMinutes);
    }

    [Fact]
    public async Task Weekly_series_expands_and_one_occurrence_can_be_deleted()
    {
        var monday = Local(2026, 11, 2, 10);
        var (series, _) = await _calendar.SaveAsync(_me, null,
            Draft("1:1", monday, TimeSpan.FromHours(1), Repeat.Weekly, until: new DateOnly(2026, 11, 30)));

        var month = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 1, 0), Local(2026, 12, 31, 0));
        Assert.Equal(5, month.Count);   // 2, 9, 16, 23, 30 November
        Assert.All(month, o => Assert.Equal(DayOfWeek.Monday, o.Start.ToLocalTime().DayOfWeek));
        Assert.All(month, o => Assert.Equal(10, o.Start.ToLocalTime().Hour));
        Assert.NotNull((await _calendar.GetAsync(series.Id))!.SeriesEndUtc);

        await _calendar.DeleteAsync(_me, series, month[1].Start, default);
        var after = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 1, 0), Local(2026, 12, 31, 0));
        Assert.Equal(4, after.Count);
        Assert.DoesNotContain(after, o => o.Start == month[1].Start);

        // Editing the series keeps the deleted occurrence deleted.
        await _calendar.SaveAsync(_me, (await _calendar.GetAsync(series.Id))!,
            Draft("1:1 renamed", monday, TimeSpan.FromHours(1), Repeat.Weekly, until: new DateOnly(2026, 11, 30)));
        var edited = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 1, 0), Local(2026, 12, 31, 0));
        Assert.Equal(4, edited.Count);
        Assert.All(edited, o => Assert.Equal("1:1 renamed", o.Appointment.Title));
    }

    [Fact]
    public async Task Endless_series_is_found_years_later()
    {
        await _calendar.SaveAsync(_me, null, Draft("Payday", Local(2026, 1, 25, 9), TimeSpan.FromHours(1), Repeat.Monthly));
        var later = await _calendar.GetOccurrencesAsync(_me.Id, Local(2030, 6, 1, 0), Local(2030, 7, 1, 0));
        Assert.Equal(25, Assert.Single(later).Start.ToLocalTime().Day);
    }

    [Fact]
    public async Task Occurrence_running_into_the_range_is_included()
    {
        await _calendar.SaveAsync(_me, null, Draft("Night shift", Local(2026, 11, 2, 22), TimeSpan.FromHours(4), Repeat.Daily));
        var morning = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 4, 0), Local(2026, 11, 4, 12));
        Assert.Equal(Local(2026, 11, 3, 22), Assert.Single(morning).Start);
    }

    [Fact]
    public async Task All_day_items_are_dates()
    {
        var day = new DateTimeOffset(2026, 12, 25, 0, 0, 0, TimeSpan.Zero);
        await _calendar.SaveAsync(_me, null, Draft("Holiday", day, TimeSpan.FromDays(1), allDay: true));
        var occurrence = Assert.Single(await _calendar.GetOccurrencesAsync(_me.Id, day.AddHours(-6), day.AddHours(30)));
        Assert.True(occurrence.IsAllDay);
        Assert.Equal(day, occurrence.Start);
        Assert.Equal(day.AddDays(1), occurrence.End);
    }

    [Fact]
    public async Task Meeting_invitation_is_sent_parsed_answered_and_recorded()
    {
        // I invite Ana.
        var start = Local(2026, 11, 5, 14);
        var (meeting, _) = await _calendar.SaveAsync(_me, null,
            Draft("Budget review", start, TimeSpan.FromHours(1), attendees: "Ana <ana@x.test>"));
        var invite = Assert.Single(_mail.Sent);
        Assert.Equal(["Ana <ana@x.test>"], invite.To);
        Assert.Equal("REQUEST", invite.Calendar!.Method);

        // Ana's client reads it...
        var parsed = _calendar.ParseInvitation(invite.Calendar.Ics)!;
        Assert.Equal("REQUEST", parsed.Method);
        Assert.Equal("Budget review", parsed.Title);
        Assert.Equal(start, parsed.Start);
        Assert.Equal("me@x.test", parsed.OrganizerEmail);
        Assert.Equal(meeting.Uid, parsed.Uid);

        // ...and accepts: it lands in her calendar and a REPLY goes back to me.
        await _calendar.RespondAsync(_ana, parsed, MeetingResponse.Accepted, sendReply: true);
        var hers = Assert.Single(await _calendar.GetOccurrencesAsync(_ana.Id, start.AddHours(-1), start.AddHours(2)));
        Assert.Equal(MeetingResponse.Accepted, hers.Appointment.MyResponse);
        Assert.False(hers.Appointment.IsOrganizer);
        var replyMail = _mail.Sent[^1];
        Assert.Equal(["me@x.test"], replyMail.To);
        Assert.Equal("Accepted: Budget review", replyMail.Subject);

        // I receive the reply: her answer is recorded on my copy.
        var reply = _calendar.ParseInvitation(replyMail.Calendar!.Ics)!;
        Assert.Equal("REPLY", reply.Method);
        Assert.Equal(MeetingResponse.Accepted, reply.ReplyStatus);
        await _calendar.ApplyAsync(_me, reply);
        var mine = (await _calendar.GetAsync(meeting.Id))!;
        Assert.Contains(CalendarService.AttendeeResponses(mine), r => r.Response == MeetingResponse.Accepted);
    }

    [Fact]
    public async Task Cancelling_a_meeting_removes_it_from_attendee_calendars()
    {
        var start = Local(2026, 11, 6, 9);
        var (meeting, _) = await _calendar.SaveAsync(_me, null, Draft("Offsite", start, TimeSpan.FromHours(2), attendees: "ana@x.test"));
        await _calendar.RespondAsync(_ana, _calendar.ParseInvitation(_mail.Sent[0].Calendar!.Ics)!, MeetingResponse.Accepted, sendReply: false);

        await _calendar.DeleteAsync(_me, meeting, null);
        var cancel = _mail.Sent[^1];
        Assert.Equal("CANCEL", cancel.Calendar!.Method);
        Assert.StartsWith("Canceled:", cancel.Subject);

        await _calendar.ApplyAsync(_ana, _calendar.ParseInvitation(cancel.Calendar.Ics)!);
        Assert.Empty(await _calendar.GetOccurrencesAsync(_ana.Id, start.AddDays(-1), start.AddDays(1)));
    }

    [Fact]
    public async Task Declining_keeps_the_meeting_out_of_the_calendar()
    {
        var start = Local(2026, 11, 7, 9);
        await _calendar.SaveAsync(_me, null, Draft("Optional thing", start, TimeSpan.FromHours(1), attendees: "ana@x.test"));
        await _calendar.RespondAsync(_ana, _calendar.ParseInvitation(_mail.Sent[0].Calendar!.Ics)!, MeetingResponse.Declined, sendReply: true);

        Assert.Empty(await _calendar.GetOccurrencesAsync(_ana.Id, start.AddDays(-1), start.AddDays(1)));
        Assert.Equal("Declined: Optional thing", _mail.Sent[^1].Subject);
    }

    [Fact]
    public async Task Moving_a_single_appointment_changes_its_time()
    {
        var start = Local(2026, 11, 2, 9);
        await _calendar.SaveAsync(_me, null, Draft("Dentist", start, TimeSpan.FromMinutes(30)));
        var occurrence = Assert.Single(await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 3, 0)));

        await _calendar.MoveAsync(_me, occurrence, Local(2026, 11, 3, 14), Local(2026, 11, 3, 15));

        Assert.Empty(await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 3, 0)));
        var moved = Assert.Single(await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 3, 0), Local(2026, 11, 4, 0)));
        Assert.Equal(Local(2026, 11, 3, 14), moved.Start);
        Assert.Equal(Local(2026, 11, 3, 15), moved.End);   // resized too: 30 min became 1 h
        Assert.Equal("Dentist", moved.Appointment.Title);
    }

    [Fact]
    public async Task Moving_one_occurrence_of_a_series_leaves_the_others()
    {
        await _calendar.SaveAsync(_me, null, Draft("Standup", Local(2026, 11, 2, 9), TimeSpan.FromMinutes(15), Repeat.Daily,
            until: new DateOnly(2026, 11, 6)));
        var week = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 7, 0));
        Assert.Equal(5, week.Count);
        var wednesday = week[2];

        // Drag Wednesday's standup to 11:00 and make it 30 minutes.
        await _calendar.MoveAsync(_me, wednesday, Local(2026, 11, 4, 11), Local(2026, 11, 4, 11, 30));
        var after = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 7, 0));
        Assert.Equal(5, after.Count);
        var moved = Assert.Single(after, o => o.IsException);
        Assert.Equal(Local(2026, 11, 4, 11), moved.Start);
        Assert.Equal(Local(2026, 11, 4, 11, 30), moved.End);
        Assert.Equal(wednesday.Start, moved.RecurrenceId);   // still known as Wednesday 9:00
        Assert.All(after.Where(o => !o.IsException), o => Assert.Equal(9, o.Start.ToLocalTime().Hour));

        // Moving it again edits the same override instead of adding another.
        await _calendar.MoveAsync(_me, moved, Local(2026, 11, 4, 16), Local(2026, 11, 4, 16, 30));
        var again = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 7, 0));
        Assert.Equal(5, again.Count);
        Assert.Equal(Local(2026, 11, 4, 16), Assert.Single(again, o => o.IsException).Start);

        // Deleting the moved occurrence removes it, not its neighbours.
        var toDelete = Assert.Single(again, o => o.IsException);
        await _calendar.DeleteAsync(_me, toDelete.Appointment, toDelete.RecurrenceId);
        var final = await _calendar.GetOccurrencesAsync(_me.Id, Local(2026, 11, 2, 0), Local(2026, 11, 7, 0));
        Assert.Equal(4, final.Count);
        Assert.DoesNotContain(final, o => o.Start.ToLocalTime().Day == 4);
    }

    [Fact]
    public async Task Moving_a_meeting_I_organise_sends_an_update()
    {
        var start = Local(2026, 11, 9, 10);
        await _calendar.SaveAsync(_me, null, Draft("Review", start, TimeSpan.FromHours(1), attendees: "ana@x.test"));
        var occurrence = Assert.Single(await _calendar.GetOccurrencesAsync(_me.Id, start.AddHours(-1), start.AddHours(2)));

        var message = await _calendar.MoveAsync(_me, occurrence, start.AddHours(2), start.AddHours(3));

        Assert.Equal("Update sent to attendees.", message);
        var update = _calendar.ParseInvitation(_mail.Sent[^1].Calendar!.Ics)!;
        Assert.Equal("REQUEST", update.Method);
        Assert.Equal(1, update.Sequence);   // attendees' calendars replace the older version
        Assert.Equal(start.AddHours(2), update.Start);
    }

    [Fact]
    public async Task Meetings_organised_by_others_cannot_be_moved()
    {
        var start = Local(2026, 11, 10, 10);
        await _calendar.SaveAsync(_me, null, Draft("Their meeting", start, TimeSpan.FromHours(1), attendees: "ana@x.test"));
        await _calendar.RespondAsync(_ana, _calendar.ParseInvitation(_mail.Sent[0].Calendar!.Ics)!, MeetingResponse.Accepted, sendReply: false);
        var hers = Assert.Single(await _calendar.GetOccurrencesAsync(_ana.Id, start.AddHours(-1), start.AddHours(2)));

        await Assert.ThrowsAsync<InvalidOperationException>(() => _calendar.MoveAsync(_ana, hers, start.AddHours(1), start.AddHours(2)));
    }

    [Fact]
    public void Broken_calendar_text_is_ignored()
    {
        Assert.Null(_calendar.ParseInvitation("not a calendar"));
    }

    [Fact]
    public void Invitation_mime_has_text_and_calendar_alternatives()
    {
        var mime = Outlook.Mail.ImapMailService.BuildMime(_me, new OutgoingMessage(["ana@x.test"], [], [], "Hi", "When: now",
            Calendar: new OutgoingCalendar("BEGIN:VCALENDAR\r\nEND:VCALENDAR\r\n", "REQUEST")));

        var alternative = Assert.IsType<MimeKit.MultipartAlternative>(mime.Body);
        var calendarPart = Assert.Single(alternative.OfType<MimeKit.TextPart>(), p => p.ContentType.IsMimeType("text", "calendar"));
        Assert.Equal("REQUEST", calendarPart.ContentType.Parameters["method"]);
        Assert.NotNull(Outlook.Mail.ImapMailService.FindCalendar(mime));
    }

    public void Dispose()
    {
        Microsoft.Data.Sqlite.SqliteConnection.ClearAllPools();
        File.Delete(_dbPath);
    }
}

public class DayLayoutTests
{
    private static (DateTimeOffset, DateTimeOffset) At(int startHour, int endHour, int startMinute = 0) =>
        (new DateTimeOffset(2026, 11, 2, startHour, startMinute, 0, TimeSpan.Zero), new DateTimeOffset(2026, 11, 2, endHour, 0, 0, TimeSpan.Zero));

    [Fact]
    public void Separate_items_get_full_width()
    {
        var slots = DayLayout.Arrange([At(9, 10), At(11, 12)]);
        Assert.All(slots, s => Assert.Equal((0, 1), (s.Lane, s.LaneCount)));
    }

    [Fact]
    public void Overlapping_items_share_the_width_of_their_cluster()
    {
        // 9-12 overlaps 10-11 and 11-13; 10-11 and 11-13 do not overlap, so they can share lane 1.
        var slots = DayLayout.Arrange([At(9, 12), At(10, 11), At(11, 13), At(14, 15)]);
        Assert.Equal((0, 2), (slots[0].Lane, slots[0].LaneCount));
        Assert.Equal((1, 2), (slots[1].Lane, slots[1].LaneCount));
        Assert.Equal((1, 2), (slots[2].Lane, slots[2].LaneCount));
        Assert.Equal((0, 1), (slots[3].Lane, slots[3].LaneCount));
    }

    [Fact]
    public void Three_way_overlap_uses_three_lanes()
    {
        var slots = DayLayout.Arrange([At(9, 11), At(9, 11, 30), At(10, 11)]);   // 9-11, 9:30-11, 10-11
        Assert.Equal(3, slots.Select(s => s.Lane).Distinct().Count());
        Assert.All(slots, s => Assert.Equal(3, s.LaneCount));
    }
}
