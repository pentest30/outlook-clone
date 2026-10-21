using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia.Media;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Calendar;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

public enum CalendarViewMode
{
    Day,
    WorkWeek,
    Week,
    Month,
}

/// <summary>An appointment drawn in a time grid column; the TimeGridPanel reads these to lay it out.</summary>
public sealed class EventBlockViewModel(AppointmentOccurrence occurrence, DateTime day, int lane, int laneCount, IBrush accent)
{
    public AppointmentOccurrence Occurrence { get; } = occurrence;
    public IBrush Accent { get; } = accent;
    public int Lane { get; } = lane;
    public int LaneCount { get; } = laneCount;

    /// <summary>Minutes from midnight of this column's day, clipped to the day (for items crossing midnight).</summary>
    public double StartMinute { get; } = Math.Max(0, (occurrence.Start.ToLocalTime().DateTime - day).TotalMinutes);
    public double EndMinute { get; } = Math.Min(24 * 60, (occurrence.End.ToLocalTime().DateTime - day).TotalMinutes);

    public string Title => Occurrence.Appointment.Title;
    public string Location => Occurrence.Appointment.Location;
    public bool HasLocation => !string.IsNullOrWhiteSpace(Location);
    public string TimeText =>
        $"{Occurrence.Start.ToLocalTime():t} – {Occurrence.End.ToLocalTime():t}";
    public bool IsTentative => Occurrence.Appointment.ShowAs == ShowAs.Tentative || Occurrence.Appointment.MyResponse == MeetingResponse.NeedsAction;
    public bool IsRecurring => Occurrence.Appointment.IsRecurring;
    public bool IsMeeting => Occurrence.Appointment.IsMeeting;
}

/// <summary>An all-day item, or a line in a month cell.</summary>
public sealed class EventChipViewModel(AppointmentOccurrence occurrence, IBrush accent)
{
    public AppointmentOccurrence Occurrence { get; } = occurrence;
    public IBrush Accent { get; } = accent;
    public string Title => Occurrence.Appointment.Title;
    public string Text => Occurrence.IsAllDay ? Title : $"{Occurrence.Start.ToLocalTime():t} {Title}";
}

public sealed partial class CalendarDayViewModel(DateTime date) : ObservableObject
{
    public DateTime Date { get; } = date;
    public string DayName => Date.ToString("ddd", CultureInfo.CurrentCulture);
    public string DayNumber => Date.Day.ToString(CultureInfo.CurrentCulture);
    public bool IsToday => Date == DateTime.Today;
    public bool IsWeekend => Date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday;

    public ObservableCollection<EventBlockViewModel> Events { get; } = [];
    public ObservableCollection<EventChipViewModel> AllDay { get; } = [];

    /// <summary>The red "now" line, in minutes from midnight; negative (hidden) except on today.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(NowTop), nameof(HasNowLine))]
    private double _nowMinute = -1;

    public double NowTop => Math.Max(0, NowMinute) / 60 * CalendarViewModel.HourHeight;
    public bool HasNowLine => NowMinute >= 0;
}

public sealed class MonthCellViewModel(DateTime date, bool inMonth, IReadOnlyList<EventChipViewModel> items)
{
    public const int VisibleItems = 3;
    public DateTime Date { get; } = date;
    public bool InMonth { get; } = inMonth;
    public bool IsToday => Date == DateTime.Today;
    public string DayNumber => Date.Day == 1 ? Date.ToString("d MMM", CultureInfo.CurrentCulture) : Date.Day.ToString(CultureInfo.CurrentCulture);
    public IReadOnlyList<EventChipViewModel> Items { get; } = items.Take(VisibleItems).ToList();
    public int MoreCount { get; } = Math.Max(0, items.Count - VisibleItems);
    public bool HasMore => MoreCount > 0;
    public string MoreText => $"+{MoreCount} more";
}

/// <summary>The Calendar module: Day, Work week, Week and Month views over the local calendar.</summary>
public partial class CalendarViewModel : ObservableObject
{
    public const double HourHeight = 48;

    private readonly ICalendarService _calendar;
    private readonly Func<Account?> _account;
    private readonly DispatcherTimer _clock;
    private int _reloadGeneration;

    public CalendarViewModel(ICalendarService calendar, Func<Account?> account)
    {
        _calendar = calendar;
        _account = account;
        // Moves the "now" line and rolls "today" over at midnight.
        _clock = new DispatcherTimer { Interval = TimeSpan.FromMinutes(1) };
        _clock.Tick += (_, _) => UpdateNowLine();
        _clock.Start();
    }

    /// <summary>Raised to open the appointment editor; the view owns windows.</summary>
    public event Action<AppointmentEditorViewModel>? EditorRequested;

    /// <summary>Set by the view: asks "this occurrence or the whole series?" (null = cancel).</summary>
    public Func<AppointmentOccurrence, Task<bool?>>? AskOccurrenceOrSeries { get; set; }

    public ObservableCollection<CalendarDayViewModel> Days { get; } = [];
    public ObservableCollection<MonthCellViewModel> MonthCells { get; } = [];
    public string[] WeekdayNames { get; private set; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsMonthView), nameof(IsTimeGridView), nameof(IsDayMode), nameof(IsWorkWeekMode), nameof(IsWeekMode), nameof(IsMonthMode))]
    private CalendarViewMode _mode = CalendarViewMode.WorkWeek;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(PickerDate))]
    private DateTime _selectedDate = DateTime.Today;

    /// <summary>The mini month calendar binds a nullable date; clearing it there keeps the current day.</summary>
    public DateTime? PickerDate
    {
        get => SelectedDate;
        set { if (value is { } date && date.Date != SelectedDate.Date) SelectedDate = date.Date; }
    }
    [ObservableProperty] private string _title = "";
    [ObservableProperty] private string _status = "";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(DeleteSelectedCommand), nameof(OpenSelectedCommand))]
    private AppointmentOccurrence? _selected;

    public bool IsMonthView => Mode == CalendarViewMode.Month;
    public bool IsTimeGridView => !IsMonthView;
    public bool IsDayMode => Mode == CalendarViewMode.Day;
    public bool IsWorkWeekMode => Mode == CalendarViewMode.WorkWeek;
    public bool IsWeekMode => Mode == CalendarViewMode.Week;
    public bool IsMonthMode => Mode == CalendarViewMode.Month;
    public double DayHeight => 24 * HourHeight;
    public IReadOnlyList<string> Hours { get; } =
        Enumerable.Range(0, 24).Select(h => DateTime.Today.AddHours(h).ToString("t", CultureInfo.CurrentCulture)).ToList();

    partial void OnModeChanged(CalendarViewMode value) => _ = ReloadAsync();
    partial void OnSelectedDateChanged(DateTime value) => _ = ReloadAsync();

    [RelayCommand] private void SetDay() => Mode = CalendarViewMode.Day;
    [RelayCommand] private void SetWorkWeek() => Mode = CalendarViewMode.WorkWeek;
    [RelayCommand] private void SetWeek() => Mode = CalendarViewMode.Week;
    [RelayCommand] private void SetMonth() => Mode = CalendarViewMode.Month;
    [RelayCommand] private void Today() => SelectedDate = DateTime.Today;

    [RelayCommand]
    private void Previous() => SelectedDate = Mode switch
    {
        CalendarViewMode.Day => SelectedDate.AddDays(-1),
        CalendarViewMode.Month => SelectedDate.AddMonths(-1),
        _ => SelectedDate.AddDays(-7),
    };

    [RelayCommand]
    private void Next() => SelectedDate = Mode switch
    {
        CalendarViewMode.Day => SelectedDate.AddDays(1),
        CalendarViewMode.Month => SelectedDate.AddMonths(1),
        _ => SelectedDate.AddDays(7),
    };

    /// <summary>Opens a day from the month view (Outlook: double-click a day number).</summary>
    public void GoToDay(DateTime date)
    {
        SelectedDate = date;
        Mode = CalendarViewMode.Day;
    }

    private static DateTime StartOfWeek(DateTime date)
    {
        var first = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
        return date.AddDays(-(((int)date.DayOfWeek - (int)first + 7) % 7)).Date;
    }

    /// <summary>The visible days for the time-grid views.</summary>
    private List<DateTime> VisibleDays()
    {
        var date = SelectedDate.Date;
        return Mode switch
        {
            CalendarViewMode.Day => [date],
            // Monday to Friday, whatever the culture's first day of week.
            CalendarViewMode.WorkWeek => Enumerable.Range(0, 5).Select(i => date.AddDays(-(((int)date.DayOfWeek + 6) % 7) + i)).ToList(),
            _ => Enumerable.Range(0, 7).Select(i => StartOfWeek(date).AddDays(i)).ToList(),
        };
    }

    public async Task ReloadAsync()
    {
        if (_account() is not { } account) return;
        var generation = ++_reloadGeneration;
        Selected = null;

        if (Mode == CalendarViewMode.Month)
        {
            var first = new DateTime(SelectedDate.Year, SelectedDate.Month, 1);
            var gridStart = StartOfWeek(first);
            var gridEnd = gridStart.AddDays(42);
            var occurrences = await LoadAsync(account, gridStart, gridEnd);
            if (generation != _reloadGeneration) return;

            Title = first.ToString("MMMM yyyy", CultureInfo.CurrentCulture);
            WeekdayNames = Enumerable.Range(0, 7).Select(i => gridStart.AddDays(i).ToString("dddd", CultureInfo.CurrentCulture)).ToArray();
            OnPropertyChanged(nameof(WeekdayNames));
            MonthCells.Clear();
            for (var i = 0; i < 42; i++)
            {
                var day = gridStart.AddDays(i);
                var items = occurrences.Where(o => Touches(o, day))
                    .OrderByDescending(o => o.IsAllDay).ThenBy(o => o.Start)
                    .Select(o => new EventChipViewModel(o, AccentFor(o.Appointment))).ToList();
                MonthCells.Add(new MonthCellViewModel(day, day.Month == first.Month, items));
            }
            return;
        }

        var days = VisibleDays();
        var all = await LoadAsync(account, days[0], days[^1].AddDays(1));
        if (generation != _reloadGeneration) return;

        Title = days.Count == 1
            ? days[0].ToString("D", CultureInfo.CurrentCulture)
            : days[0].Month == days[^1].Month
                ? $"{days[0].Day}–{days[^1].ToString("d MMMM yyyy", CultureInfo.CurrentCulture)}"
                : $"{days[0].ToString("d MMM", CultureInfo.CurrentCulture)} – {days[^1].ToString("d MMM yyyy", CultureInfo.CurrentCulture)}";

        Days.Clear();
        foreach (var date in days)
        {
            var day = new CalendarDayViewModel(date);
            var dayItems = all.Where(o => Touches(o, date)).ToList();
            foreach (var o in dayItems.Where(o => o.IsAllDay))
                day.AllDay.Add(new EventChipViewModel(o, AccentFor(o.Appointment)));

            var timed = dayItems.Where(o => !o.IsAllDay).ToList();
            var slots = DayLayout.Arrange(timed.Select(o => (o.Start, o.End)).ToList());
            foreach (var slot in slots)
                day.Events.Add(new EventBlockViewModel(timed[slot.Index], date, slot.Lane, slot.LaneCount, AccentFor(timed[slot.Index].Appointment)));
            Days.Add(day);
        }
        UpdateNowLine();
    }

    private async Task<IReadOnlyList<AppointmentOccurrence>> LoadAsync(Account account, DateTime from, DateTime to)
    {
        try
        {
            Status = "";
            // All-day items are floating dates stored at 00:00 UTC, so widen the query by a day each side.
            var occurrences = await _calendar.GetOccurrencesAsync(account.Id,
                new DateTimeOffset(from.AddDays(-1)), new DateTimeOffset(to.AddDays(1)));
            return occurrences;
        }
        catch (Exception ex)
        {
            Status = $"Couldn't load the calendar: {ex.Message}";
            return [];
        }
    }

    /// <summary>Whether an occurrence shows on a given local day.</summary>
    private static bool Touches(AppointmentOccurrence o, DateTime day)
    {
        if (o.IsAllDay)
        {
            var firstDay = o.Start.UtcDateTime.Date;
            var lastDay = o.End.UtcDateTime.Date.AddDays(-1);   // end date is exclusive
            return day >= firstDay && day <= (lastDay < firstDay ? firstDay : lastDay);
        }
        var start = o.Start.ToLocalTime().DateTime;
        var end = o.End.ToLocalTime().DateTime;
        return start < day.AddDays(1) && (end > day || (end == start && start.Date == day));
    }

    /// <summary>Outlook colours: blue for my items, purple for meetings, hatched-looking grey for tentative ones.</summary>
    private static IBrush AccentFor(Appointment a) =>
        a.MyResponse == MeetingResponse.NeedsAction || a.ShowAs == ShowAs.Tentative ? Brushes.Gray
        : a.IsMeeting ? new SolidColorBrush(Color.Parse("#8764B8"))
        : a.ShowAs == ShowAs.OutOfOffice ? new SolidColorBrush(Color.Parse("#B4009E"))
        : new SolidColorBrush(Color.Parse("#0F6CBD"));

    private void UpdateNowLine()
    {
        var now = DateTime.Now;
        foreach (var day in Days)
            day.NowMinute = day.IsToday ? (now - now.Date).TotalMinutes : -1;
    }

    // ───────────────────────────── Create / edit / delete ─────────────────────────────

    [RelayCommand]
    private void NewAppointment()
    {
        // Next half hour on the selected day, like Outlook's New Appointment.
        var now = DateTime.Now;
        var start = SelectedDate.Date == now.Date
            ? now.Date.AddMinutes(Math.Ceiling((now - now.Date).TotalMinutes / 30) * 30)
            : SelectedDate.Date.AddHours(9);
        NewAppointmentAt(start, allDay: false);
    }

    public void NewAppointmentAt(DateTime start, bool allDay)
    {
        if (_account() is not { } account) return;
        var editor = new AppointmentEditorViewModel(_calendar, account, null, null);
        editor.Initialize(start, allDay);
        Open(editor);
    }

    public async Task OpenAsync(AppointmentOccurrence occurrence)
    {
        if (_account() is not { } account) return;
        var appointment = await _calendar.GetAsync(occurrence.Appointment.Id);
        if (appointment is null)
        {
            await ReloadAsync();
            return;
        }
        Open(new AppointmentEditorViewModel(_calendar, account, appointment, _calendar.ToDraft(appointment)));
    }

    private void Open(AppointmentEditorViewModel editor)
    {
        editor.Saved += message => Dispatcher.UIThread.Post(() =>
        {
            Status = message ?? "";
            _ = ReloadAsync();
        });
        EditorRequested?.Invoke(editor);
    }

    private bool HasSelected() => Selected is not null;

    [RelayCommand(CanExecute = nameof(HasSelected))]
    private Task OpenSelectedAsync() => Selected is { } s ? OpenAsync(s) : Task.CompletedTask;

    [RelayCommand(CanExecute = nameof(HasSelected))]
    private Task DeleteSelectedAsync() => Selected is { } s ? DeleteAsync(s) : Task.CompletedTask;

    public async Task DeleteAsync(AppointmentOccurrence occurrence)
    {
        if (_account() is not { } account) return;
        DateTimeOffset? onlyThis = null;
        if (occurrence.Appointment.IsRecurring)
        {
            var answer = AskOccurrenceOrSeries is null ? false : await AskOccurrenceOrSeries(occurrence);
            if (answer is null) return;
            if (answer == true) onlyThis = occurrence.Start;
        }
        try
        {
            Status = await _calendar.DeleteAsync(account, occurrence.Appointment, onlyThis) ?? "";
        }
        catch (Exception ex)
        {
            Status = $"Couldn't delete: {ex.Message}";
        }
        await ReloadAsync();
    }
}
