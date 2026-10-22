using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Calendar;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

/// <summary>The appointment / meeting window. Adding attendees turns it into a meeting invitation.</summary>
public partial class AppointmentEditorViewModel : ObservableObject
{
    private readonly ICalendarService _calendar;
    private readonly Account _account;
    private readonly Appointment? _existing;

    public AppointmentEditorViewModel(ICalendarService calendar, Account account, Appointment? existing, AppointmentDraft? draft)
    {
        _calendar = calendar;
        _account = account;
        _existing = existing;
        if (draft is not null)
            Load(draft);
        if (existing is not null && existing.IsOrganizer && existing.IsMeeting)
            Responses = string.Join(", ", CalendarService.AttendeeResponses(existing)
                .Select(r => $"{r.Who}: {r.Response switch { MeetingResponse.Accepted => "accepted", MeetingResponse.Tentative => "tentative", MeetingResponse.Declined => "declined", _ => "no response" }}"));
    }

    /// <summary>Raised after saving or deleting, with a message worth showing (or null).</summary>
    public event Action<string?>? Saved;
    public event Action? CloseRequested;

    public static Repeat[] RepeatOptions { get; } = Enum.GetValues<Repeat>();
    public static ShowAs[] ShowAsOptions { get; } = Enum.GetValues<ShowAs>();

    /// <summary>Outlook's reminder choices, in minutes (null = none).</summary>
    public static ReminderOption[] ReminderOptions { get; } =
    [
        new(null, "None"), new(0, "At start time"), new(5, "5 minutes"), new(15, "15 minutes"), new(30, "30 minutes"),
        new(60, "1 hour"), new(120, "2 hours"), new(1440, "1 day"), new(10080, "1 week"),
    ];

    /// <summary>Address suggestions for the recipient fields (set by the main window).</summary>
    public Func<string, CancellationToken, Task<IEnumerable<object>>>? SuggestRecipients { get; set; }

    public bool IsNew => _existing is null;
    /// <summary>Invitations from other people are read-only here; answer them from the e-mail.</summary>
    public bool CanEdit => _existing is null || _existing.IsOrganizer;
    public string OrganizerText => _existing is { IsOrganizer: false } a
        ? $"Organizer: {(string.IsNullOrWhiteSpace(a.OrganizerName) ? a.OrganizerEmail : a.OrganizerName)} · Your response: {a.MyResponse}"
        : "";
    public bool HasOrganizerText => OrganizerText.Length > 0;

    [ObservableProperty] [NotifyPropertyChangedFor(nameof(WindowTitle))] private string _subject = "";
    [ObservableProperty] private string _location = "";
    [ObservableProperty] private DateTime? _startDate = DateTime.Today;
    [ObservableProperty] private TimeSpan? _startTime = TimeSpan.FromHours(9);
    [ObservableProperty] private DateTime? _endDate = DateTime.Today;
    [ObservableProperty] private TimeSpan? _endTime = TimeSpan.FromHours(9.5);
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(ShowTimes))] private bool _isAllDay;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(ShowRepeatUntil))] private Repeat _repeat;
    [ObservableProperty] private DateTime? _repeatUntil;
    [ObservableProperty] private ReminderOption _reminder = ReminderOptions[3];
    [ObservableProperty] private ShowAs _showAs;
    [ObservableProperty] [NotifyPropertyChangedFor(nameof(SaveLabel))] private string _attendees = "";
    [ObservableProperty] private string _notes = "";
    [ObservableProperty] private string _error = "";
    [ObservableProperty] private string _responses = "";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SaveCommand), nameof(DeleteCommand))]
    private bool _isBusy;

    public bool ShowTimes => !IsAllDay;
    public bool ShowRepeatUntil => Repeat != Repeat.Never;
    public bool HasResponses => Responses.Length > 0;
    public string SaveLabel => string.IsNullOrWhiteSpace(Attendees) ? "Save & close" : IsNew ? "Send" : "Send update";
    public string WindowTitle => (string.IsNullOrWhiteSpace(Subject) ? "Untitled" : Subject) + (string.IsNullOrWhiteSpace(Attendees) ? " - Appointment" : " - Meeting");

    private TimeSpan _length = TimeSpan.FromMinutes(30);
    private bool _syncingEnd;

    public void Initialize(DateTime start, bool allDay, int? reminderMinutes = 15)
    {
        Reminder = ReminderOptions.FirstOrDefault(r => r.Minutes == reminderMinutes) ?? ReminderOptions[3];
        IsAllDay = allDay;
        StartDate = start.Date;
        StartTime = start.TimeOfDay;
        var end = allDay ? start.Date : start.AddMinutes(30);
        EndDate = end.Date;
        EndTime = end.TimeOfDay;
    }

    private void Load(AppointmentDraft draft)
    {
        Subject = draft.Title;
        Location = draft.Location;
        IsAllDay = draft.IsAllDay;
        if (draft.IsAllDay)
        {
            // Stored end is exclusive (the next day); the editor shows the last day, like Outlook.
            StartDate = draft.Start.UtcDateTime.Date;
            EndDate = draft.End.UtcDateTime.Date.AddDays(-1);
        }
        else
        {
            var start = draft.Start.ToLocalTime();
            var end = draft.End.ToLocalTime();
            StartDate = start.Date;
            StartTime = start.TimeOfDay;
            EndDate = end.Date;
            EndTime = end.TimeOfDay;
        }
        Repeat = draft.Repeat;
        RepeatUntil = draft.RepeatUntil?.ToDateTime(TimeOnly.MinValue);
        Reminder = ReminderOptions.FirstOrDefault(r => r.Minutes == draft.ReminderMinutes) ?? new ReminderOption(draft.ReminderMinutes, $"{draft.ReminderMinutes} minutes");
        ShowAs = draft.ShowAs;
        Attendees = draft.Attendees;
        Notes = draft.Notes;
        _length = (EndAt() ?? StartAt()!.Value) - StartAt()!.Value;
    }

    // Moving the start keeps the duration, as Outlook does.
    partial void OnStartDateChanged(DateTime? value) => KeepLength();
    partial void OnStartTimeChanged(TimeSpan? value) => KeepLength();
    partial void OnEndDateChanged(DateTime? value) => RememberLength();
    partial void OnEndTimeChanged(TimeSpan? value) => RememberLength();

    private void KeepLength()
    {
        if (StartAt() is not { } start || _syncingEnd) return;
        _syncingEnd = true;
        var end = start + _length;
        EndDate = end.Date;
        EndTime = end.TimeOfDay;
        _syncingEnd = false;
    }

    private void RememberLength()
    {
        if (_syncingEnd || StartAt() is not { } start || EndAt() is not { } end) return;
        if (end > start) _length = end - start;
    }

    private DateTime? StartAt() => StartDate is { } d ? d.Date + (IsAllDay ? TimeSpan.Zero : StartTime ?? TimeSpan.Zero) : null;
    private DateTime? EndAt() => EndDate is { } d ? d.Date + (IsAllDay ? TimeSpan.Zero : EndTime ?? TimeSpan.Zero) : null;

    private bool CanSave() => !IsBusy && CanEdit;

    [RelayCommand(CanExecute = nameof(CanSave))]
    private async Task SaveAsync()
    {
        Error = "";
        if (StartAt() is not { } start || EndAt() is not { } end)
        {
            Error = "Choose a start and an end.";
            return;
        }
        DateTimeOffset startAt, endAt;
        if (IsAllDay)
        {
            // All-day: floating dates at 00:00 UTC, end exclusive (the day after the last day).
            // (Unspecified kind: a Local DateTime, e.g. from DateTime.Today, refuses a zero offset.)
            startAt = new DateTimeOffset(DateTime.SpecifyKind(start.Date, DateTimeKind.Unspecified), TimeSpan.Zero);
            endAt = new DateTimeOffset(DateTime.SpecifyKind(end.Date.AddDays(1), DateTimeKind.Unspecified), TimeSpan.Zero);
        }
        else
        {
            startAt = new DateTimeOffset(DateTime.SpecifyKind(start, DateTimeKind.Local));
            endAt = new DateTimeOffset(DateTime.SpecifyKind(end, DateTimeKind.Local));
        }
        if (endAt <= startAt)
        {
            Error = "The end must be after the start.";
            return;
        }

        var draft = new AppointmentDraft(Subject.Trim(), Location.Trim(), startAt, endAt, IsAllDay, Repeat,
            Repeat == Repeat.Never || RepeatUntil is null ? null : DateOnly.FromDateTime(RepeatUntil.Value.Date),
            Reminder.Minutes, ShowAs, Attendees, Notes);

        IsBusy = true;
        try
        {
            var (_, problem) = await _calendar.SaveAsync(_account, _existing, draft);
            Saved?.Invoke(problem ?? (string.IsNullOrWhiteSpace(Attendees) ? null : IsNew ? "Invitation sent." : "Update sent."));
            CloseRequested?.Invoke();
        }
        catch (Exception ex)
        {
            Error = $"Couldn't save: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }

    private bool CanDelete() => !IsBusy && _existing is not null;

    /// <summary>Deletes the whole item (single occurrences are deleted from the calendar view).</summary>
    [RelayCommand(CanExecute = nameof(CanDelete))]
    private async Task DeleteAsync()
    {
        if (_existing is null) return;
        IsBusy = true;
        try
        {
            var problem = await _calendar.DeleteAsync(_account, _existing, null);
            Saved?.Invoke(problem ?? (_existing.IsOrganizer && _existing.IsMeeting ? "Cancellation sent." : null));
            CloseRequested?.Invoke();
        }
        catch (Exception ex)
        {
            Error = $"Couldn't delete: {ex.Message}";
        }
        finally
        {
            IsBusy = false;
        }
    }
}

public sealed record ReminderOption(int? Minutes, string Label)
{
    public override string ToString() => Label;
}
