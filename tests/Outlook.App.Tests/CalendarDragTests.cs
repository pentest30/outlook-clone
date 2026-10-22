using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Models;
using Outlook.Core.Services;

[assembly: AvaloniaTestApplication(typeof(Outlook.App.Tests.TestApp))]

namespace Outlook.App.Tests;

/// <summary>Runs the real App (theme and styles) on Avalonia's headless platform, rendering with Skia.</summary>
public static class TestApp
{
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>().UseSkia().UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false });
}

/// <summary>A calendar with fixed appointments that records the moves the view asks for.</summary>
internal sealed class FakeCalendar : ICalendarService
{
    public List<AppointmentOccurrence> Occurrences { get; } = [];
    public List<(AppointmentOccurrence Occurrence, DateTimeOffset Start, DateTimeOffset End)> Moves { get; } = [];

    public Task<IReadOnlyList<AppointmentOccurrence>> GetOccurrencesAsync(int accountId, DateTimeOffset from, DateTimeOffset to, CancellationToken ct = default) =>
        Task.FromResult<IReadOnlyList<AppointmentOccurrence>>(Occurrences.Where(o => o.Start < to && o.End > from).ToList());

    public Task<string?> MoveAsync(Account account, AppointmentOccurrence occurrence, DateTimeOffset newStart, DateTimeOffset newEnd, CancellationToken ct = default)
    {
        Moves.Add((occurrence, newStart, newEnd));
        return Task.FromResult<string?>(null);
    }

    public Task<Appointment?> GetAsync(int id, CancellationToken ct = default) => Task.FromResult<Appointment?>(null);
    public AppointmentDraft ToDraft(Appointment a) => throw new NotSupportedException();
    public Task<(Appointment Saved, string? Problem)> SaveAsync(Account a, Appointment? e, AppointmentDraft d, CancellationToken ct = default) => throw new NotSupportedException();
    public Task<string?> DeleteAsync(Account a, Appointment ap, DateTimeOffset? id, CancellationToken ct = default) => throw new NotSupportedException();
    public Invitation? ParseInvitation(string ics) => null;
    public Task<Appointment?> FindAsync(int accountId, string uid, CancellationToken ct = default) => Task.FromResult<Appointment?>(null);
    public Task<string?> RespondAsync(Account a, Invitation i, MeetingResponse r, bool s, CancellationToken ct = default) => throw new NotSupportedException();
    public Task ApplyAsync(Account a, Invitation i, CancellationToken ct = default) => throw new NotSupportedException();
    public Task DismissReminderAsync(Appointment a, DateTimeOffset s, CancellationToken ct = default) => Task.CompletedTask;
}

public class CalendarDragTests
{
    // A Wednesday, so the work week shown is Monday 2 – Friday 6 November.
    private static readonly DateTime Wednesday = new(2026, 11, 4);
    private static readonly string Shots = Path.Combine(Path.GetTempPath(), "outlook-ui-shots");

    private static DateTimeOffset Local(DateTime day, double hour) =>
        new(DateTime.SpecifyKind(day.AddHours(hour), DateTimeKind.Local));

    private static AppointmentOccurrence Occurrence(string title, DateTimeOffset start, DateTimeOffset end, bool organizer = true, bool allDay = false) =>
        new(new Appointment { Id = title.GetHashCode(), Title = title, StartUtc = start, EndUtc = end, IsOrganizer = organizer, IsAllDay = allDay }, start, end);

    private static (Window Window, CalendarViewModel Vm, FakeCalendar Calendar) Show(CalendarViewMode mode, params AppointmentOccurrence[] items)
    {
        var calendar = new FakeCalendar();
        calendar.Occurrences.AddRange(items);
        var account = new Account { Id = 1, Email = "me@x.test" };
        var vm = new CalendarViewModel(calendar, () => account) { SelectedDate = Wednesday, Mode = mode };
        var window = new Window { Width = 1300, Height = 900, Content = new CalendarView { DataContext = vm } };
        window.Show();
        vm.ReloadAsync().GetAwaiter().GetResult();
        Pump();
        return (window, vm, calendar);
    }

    /// <summary>Lets bindings, layout and async continuations settle.</summary>
    private static void Pump()
    {
        for (var i = 0; i < 5; i++)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        }
    }

    /// <summary>The appointment's own Border (the one with the pointer handlers), found by title.</summary>
    private static Border Block(Window window, string title) =>
        window.GetVisualDescendants().OfType<Border>()
            .Where(b => b.DataContext is EventBlockViewModel block && block.Title == title)
            .OrderByDescending(b => b.Bounds.Height)
            .First();

    private static Border Chip(Window window, string title) =>
        window.GetVisualDescendants().OfType<Border>()
            .Where(b => b.DataContext is EventChipViewModel chip && chip.Title == title && b.Padding.Left > 0)
            .First();

    private static Point At(Visual control, Window window, double x, double y) =>
        control.TranslatePoint(new Point(x, y), window) ?? throw new InvalidOperationException("Not on screen.");

    private static void Drag(Window window, Point from, Point to, string? shot = null)
    {
        window.MouseDown(from, MouseButton.Left);
        Pump();
        const int steps = 8;
        for (var i = 1; i <= steps; i++)
        {
            window.MouseMove(from + (to - from) * i / steps, RawInputModifiers.LeftMouseButton);
            Pump();
        }
        if (shot is not null) Capture(window, shot);
        window.MouseUp(to, MouseButton.Left);
        Pump();
    }

    private static void Capture(Window window, string name)
    {
        Directory.CreateDirectory(Shots);
        window.CaptureRenderedFrame()?.Save(Path.Combine(Shots, name + ".png"));
    }

    [AvaloniaFact]
    public void Dragging_the_body_moves_to_another_day_and_time()
    {
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Dentist", Local(Wednesday, 10), Local(Wednesday, 11)));
        var block = Block(window, "Dentist");
        var start = At(block, window, block.Bounds.Width / 2, block.Bounds.Height / 2);
        var columnWidth = block.Bounds.Width + 8;   // block width plus the column's gaps and margins

        // One column right (Thursday) and two hours down.
        Drag(window, start, start + new Point(columnWidth, 2 * CalendarViewModel.HourHeight), shot: "move-mid");

        var move = Assert.Single(calendar.Moves);
        Assert.Equal(Local(Wednesday.AddDays(1), 12), move.Start);
        Assert.Equal(Local(Wednesday.AddDays(1), 13), move.End);   // length kept
    }

    [AvaloniaFact]
    public void Dragging_the_bottom_edge_changes_the_end_only()
    {
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Review", Local(Wednesday, 9), Local(Wednesday, 10)));
        var block = Block(window, "Review");
        var bottomEdge = At(block, window, block.Bounds.Width / 2, block.Bounds.Height - 2);

        // 1 h 30 down, snapped to the quarter hour.
        Drag(window, bottomEdge, bottomEdge + new Point(0, 1.5 * CalendarViewModel.HourHeight), shot: "resize-mid");

        var move = Assert.Single(calendar.Moves);
        Assert.Equal(Local(Wednesday, 9), move.Start);
        Assert.Equal(Local(Wednesday, 11.5), move.End);
    }

    [AvaloniaFact]
    public void Dragging_the_top_edge_changes_the_start_only_and_keeps_a_minimum_length()
    {
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Call", Local(Wednesday, 14), Local(Wednesday, 15)));
        var block = Block(window, "Call");
        var topEdge = At(block, window, block.Bounds.Width / 2, 2);

        // Way past the end: the start stops 15 minutes before it.
        Drag(window, topEdge, topEdge + new Point(0, 3 * CalendarViewModel.HourHeight));

        var move = Assert.Single(calendar.Moves);
        Assert.Equal(Local(Wednesday, 14.75), move.Start);
        Assert.Equal(Local(Wednesday, 15), move.End);
    }

    [AvaloniaFact]
    public void Times_snap_to_the_quarter_hour()
    {
        var (window, _, calendar) = Show(CalendarViewMode.Day, Occurrence("Sync", Local(Wednesday, 10), Local(Wednesday, 10.5)));
        var block = Block(window, "Sync");
        var start = At(block, window, block.Bounds.Width / 2, block.Bounds.Height / 2);

        // 22 minutes down rounds to 15.
        Drag(window, start, start + new Point(0, 22.0 / 60 * CalendarViewModel.HourHeight));

        Assert.Equal(Local(Wednesday, 10.25), Assert.Single(calendar.Moves).Start);
    }

    [AvaloniaFact]
    public void A_click_without_movement_moves_nothing()
    {
        var (window, vm, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Lunch", Local(Wednesday, 12), Local(Wednesday, 13)));
        var block = Block(window, "Lunch");
        var point = At(block, window, block.Bounds.Width / 2, block.Bounds.Height / 2);

        window.MouseDown(point, MouseButton.Left);
        window.MouseMove(point + new Point(2, 1), RawInputModifiers.LeftMouseButton);   // under the drag threshold
        window.MouseUp(point, MouseButton.Left);
        Pump();

        Assert.Empty(calendar.Moves);
        Assert.Equal("Lunch", vm.Selected?.Appointment.Title);   // but it is selected
    }

    [AvaloniaFact]
    public void Escape_cancels_a_drag()
    {
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Standup", Local(Wednesday, 9), Local(Wednesday, 9.25)));
        var block = Block(window, "Standup");
        var start = At(block, window, block.Bounds.Width / 2, block.Bounds.Height / 2);

        window.MouseDown(start, MouseButton.Left);
        window.MouseMove(start + new Point(0, 100), RawInputModifiers.LeftMouseButton);
        Pump();
        window.KeyPressQwerty(PhysicalKey.Escape, RawInputModifiers.None);
        window.MouseUp(start + new Point(0, 100), MouseButton.Left);
        Pump();

        Assert.Empty(calendar.Moves);
        Assert.Equal(1, block.Opacity);   // the dimmed original is restored
    }

    [AvaloniaFact]
    public void Someone_elses_meeting_cannot_be_dragged()
    {
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Their meeting", Local(Wednesday, 11), Local(Wednesday, 12), organizer: false));
        var block = Block(window, "Their meeting");
        var start = At(block, window, block.Bounds.Width / 2, block.Bounds.Height / 2);

        Drag(window, start, start + new Point(0, 2 * CalendarViewModel.HourHeight));

        Assert.Empty(calendar.Moves);
    }

    [AvaloniaFact]
    public void Month_view_items_move_to_another_day_keeping_their_time()
    {
        var (window, _, calendar) = Show(CalendarViewMode.Month, Occurrence("Payroll", Local(Wednesday, 16), Local(Wednesday, 17)));
        var chip = Chip(window, "Payroll");
        var start = At(chip, window, chip.Bounds.Width / 2, chip.Bounds.Height / 2);
        var grid = window.GetVisualDescendants().OfType<ItemsControl>().First(c => c.Name == "MonthGrid");
        var cell = new Point(grid.Bounds.Width / 7, grid.Bounds.Height / 6);

        // One week later (one row down) and one day later (one column right).
        Drag(window, start, start + new Point(cell.X, cell.Y), shot: "month-mid");

        var move = Assert.Single(calendar.Moves);
        Assert.Equal(Local(Wednesday.AddDays(8), 16), move.Start);
        Assert.Equal(Local(Wednesday.AddDays(8), 17), move.End);
    }

    [AvaloniaFact]
    public void All_day_items_move_by_whole_days()
    {
        var day = new DateTimeOffset(DateTime.SpecifyKind(Wednesday, DateTimeKind.Unspecified), TimeSpan.Zero);
        var (window, _, calendar) = Show(CalendarViewMode.WorkWeek, Occurrence("Offsite", day, day.AddDays(1), allDay: true));
        var chip = Chip(window, "Offsite");
        var start = At(chip, window, chip.Bounds.Width / 2, chip.Bounds.Height / 2);
        var strip = window.GetVisualDescendants().OfType<ItemsControl>().First(c => c.Name == "AllDayStrip");

        // Two columns left: Monday.
        Drag(window, start, start - new Point(2 * strip.Bounds.Width / 5, 0));

        var move = Assert.Single(calendar.Moves);
        Assert.Equal(day.AddDays(-2), move.Start);
        Assert.Equal(day.AddDays(-1), move.End);
    }
}
