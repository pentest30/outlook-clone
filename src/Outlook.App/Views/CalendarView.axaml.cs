using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.VisualTree;
using Outlook.App.ViewModels;
using Outlook.Core.Models;

namespace Outlook.App.Views;

public partial class CalendarView : UserControl
{
    /// <summary>Times snap to quarter hours while dragging, as in Outlook.</summary>
    private static readonly TimeSpan Snap = TimeSpan.FromMinutes(15);
    /// <summary>The shortest an appointment can be resized to.</summary>
    private static readonly TimeSpan MinimumLength = TimeSpan.FromMinutes(15);
    /// <summary>Pixels from the top or bottom edge of a block that resize instead of move.</summary>
    private const double EdgeSize = 6;
    /// <summary>A press only becomes a drag after this much movement, so clicks stay clicks.</summary>
    private const double DragThreshold = 4;

    private static readonly Cursor ResizeCursor = new(StandardCursorType.SizeNorthSouth);
    private static readonly Cursor MoveCursor = new(StandardCursorType.SizeAll);

    private enum DragKind { Move, ResizeStart, ResizeEnd, MoveDays }

    /// <summary>An appointment being dragged: what, how, from where, and where it would land now.</summary>
    private sealed class DragState
    {
        public required Control Source { get; init; }
        public required IPointer Pointer { get; init; }
        public required AppointmentOccurrence Occurrence { get; init; }
        public required DragKind Kind { get; init; }
        public required Visual Surface { get; init; }
        public required Point Origin { get; init; }
        /// <summary>The days laid out on the surface, left to right, top to bottom.</summary>
        public required IReadOnlyList<DateTime> Days { get; init; }
        public required int Columns { get; init; }
        public required int Rows { get; init; }
        public required int StartIndex { get; init; }
        public bool Active { get; set; }
        public DateTimeOffset NewStart { get; set; }
        public DateTimeOffset NewEnd { get; set; }
    }

    private DragState? _drag;

    public CalendarView()
    {
        InitializeComponent();
        AddHandler(KeyDownEvent, OnKeyDownTunnel, RoutingStrategies.Tunnel);
    }

    private CalendarViewModel? Vm => DataContext as CalendarViewModel;

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (Vm is not { } vm) return;
        vm.EditorRequested += editor =>
        {
            var window = new AppointmentWindow { DataContext = editor };
            if (TopLevel.GetTopLevel(this) is Window owner) window.Show(owner);
            else window.Show();
        };
        vm.AskOccurrenceOrSeries = AskOccurrenceOrSeriesAsync;
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        // Open on the working day (8:00), as Outlook does.
        TimeScroll.Offset = new Vector(0, 8 * CalendarViewModel.HourHeight);
    }

    // ───────────────────────────── Time grid: create and open ─────────────────────────────

    /// <summary>Double-click an empty slot: new appointment at that half hour.</summary>
    private void OnDayColumnDoubleTapped(object? sender, TappedEventArgs e)
    {
        if (sender is not Control { DataContext: CalendarDayViewModel day } column || Vm is not { } vm) return;
        var minutes = e.GetPosition(column).Y / CalendarViewModel.HourHeight * 60;
        var slot = Math.Clamp(Math.Floor(minutes / 30) * 30, 0, 23.5 * 60);
        vm.NewAppointmentAt(day.Date.AddMinutes(slot), allDay: false);
    }

    private void OnAllDayDoubleTapped(object? sender, TappedEventArgs e)
    {
        if (sender is Control { DataContext: CalendarDayViewModel day } && Vm is { } vm)
            vm.NewAppointmentAt(day.Date, allDay: true);
    }

    private async void OnBlockDoubleTapped(object? sender, TappedEventArgs e)
    {
        e.Handled = true;   // not a double-click on the empty column underneath
        if (sender is Control { DataContext: EventBlockViewModel block } && Vm is { } vm)
            await vm.OpenAsync(block.Occurrence);
    }

    // ───────────────────────────── Time grid: drag to move and resize ─────────────────────────────

    private void OnBlockPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control { DataContext: EventBlockViewModel block } control || Vm is not { } vm) return;
        vm.Selected = block.Occurrence;
        Focus();   // so Delete / Enter / Escape reach the calendar
        e.Handled = true;

        var point = e.GetCurrentPoint(control);
        if (point.Properties.IsRightButtonPressed)
        {
            ShowMenu(control, block.Occurrence);
            return;
        }
        if (!point.Properties.IsLeftButtonPressed || e.ClickCount > 1 || !CalendarViewModel.CanMove(block.Occurrence))
            return;

        var day = vm.Days.FirstOrDefault(d => d.Events.Contains(block));
        if (day is null) return;
        _drag = new DragState
        {
            Source = control,
            Pointer = e.Pointer,
            Occurrence = block.Occurrence,
            Kind = EdgeAt(control, point.Position) switch
            {
                -1 => DragKind.ResizeStart,
                1 => DragKind.ResizeEnd,
                _ => DragKind.Move,
            },
            Surface = DayColumns,
            Origin = e.GetPosition(DayColumns),
            Days = vm.Days.Select(d => d.Date).ToList(),
            Columns = vm.Days.Count,
            Rows = 1,
            StartIndex = vm.Days.IndexOf(day),
        };
        e.Pointer.Capture(control);
    }

    /// <summary>-1 near the top edge, 1 near the bottom edge, 0 in the body.</summary>
    private static int EdgeAt(Control block, Point position)
    {
        var height = block.Bounds.Height;
        if (height < EdgeSize * 3) return position.Y > height / 2 ? 1 : 0;   // tiny block: bottom half resizes
        if (position.Y <= EdgeSize) return -1;
        if (position.Y >= height - EdgeSize) return 1;
        return 0;
    }

    private void OnBlockPointerMoved(object? sender, PointerEventArgs e)
    {
        if (sender is not Control { DataContext: EventBlockViewModel block } control) return;

        if (_drag is null)
        {
            // Hover feedback: a resize cursor on the edges, a move cursor on the body.
            control.Cursor = CalendarViewModel.CanMove(block.Occurrence)
                ? EdgeAt(control, e.GetPosition(control)) != 0 ? ResizeCursor : MoveCursor
                : Cursor.Default;
            return;
        }
        if (!ReferenceEquals(_drag.Source, control)) return;

        var position = e.GetPosition(DayColumns);
        if (!_drag.Active)
        {
            var moved = position - _drag.Origin;
            if (Math.Abs(moved.X) < DragThreshold && Math.Abs(moved.Y) < DragThreshold) return;
            _drag.Active = true;
            control.Opacity = 0.45;
        }
        AutoScroll(e.GetPosition(TimeScroll));
        UpdateTimeDrag(_drag, position);
    }

    /// <summary>Works out the new start and end from the pointer, snapped to the quarter hour, and shows the ghost.</summary>
    private void UpdateTimeDrag(DragState drag, Point position)
    {
        var o = drag.Occurrence;
        var start = o.Start.ToLocalTime().DateTime;
        var end = o.End.ToLocalTime().DateTime;
        var shift = TimeSpan.FromMinutes((position.Y - drag.Origin.Y) / CalendarViewModel.HourHeight * 60);
        var columnWidth = DayColumns.Bounds.Width / drag.Columns;
        var column = drag.StartIndex;

        DateTime newStart = start, newEnd = end;
        switch (drag.Kind)
        {
            case DragKind.Move:
                column = Math.Clamp((int)Math.Floor(position.X / columnWidth), 0, drag.Columns - 1);
                newStart = Round(start + (drag.Days[column] - drag.Days[drag.StartIndex]) + shift);
                newEnd = newStart + (end - start);
                break;
            case DragKind.ResizeStart:
                newStart = Min(Round(start + shift), end - MinimumLength);
                break;
            case DragKind.ResizeEnd:
                newEnd = Max(Round(end + shift), start + MinimumLength);
                break;
        }
        drag.NewStart = ToOffset(newStart);
        drag.NewEnd = ToOffset(newEnd);

        // The ghost sits in the column where the item starts, clipped to that day.
        var day = drag.Days[column];
        var top = Math.Max(0, (newStart - day).TotalMinutes) / 60 * CalendarViewModel.HourHeight;
        var bottom = Math.Min(24 * 60, (newEnd - day).TotalMinutes) / 60 * CalendarViewModel.HourHeight;
        Canvas.SetLeft(TimeGhost, column * columnWidth + 2);
        Canvas.SetTop(TimeGhost, top);
        TimeGhost.Width = Math.Max(0, columnWidth - 6);
        TimeGhost.Height = Math.Max(18, bottom - top);
        TimeGhostText.Text = newStart.Date == start.Date
            ? $"{newStart:t} – {newEnd:t}"
            : $"{newStart:ddd} {newStart:t} – {newEnd:t}";
        TimeGhost.IsVisible = true;
    }

    /// <summary>Scroll the time grid when dragging near its top or bottom edge.</summary>
    private void AutoScroll(Point inViewport)
    {
        const double zone = 32, step = 16;
        var offset = TimeScroll.Offset;
        if (inViewport.Y < zone)
            TimeScroll.Offset = offset.WithY(Math.Max(0, offset.Y - step));
        else if (inViewport.Y > TimeScroll.Bounds.Height - zone)
            TimeScroll.Offset = offset.WithY(offset.Y + step);
    }

    private async void OnBlockReleased(object? sender, PointerReleasedEventArgs e) => await FinishDragAsync();

    // ───────────────────────────── All-day strip and month view: drag to another day ─────────────────────────────

    private void OnChipPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control { DataContext: EventChipViewModel chip } control || Vm is not { } vm) return;
        vm.Selected = chip.Occurrence;
        Focus();
        e.Handled = true;

        var point = e.GetCurrentPoint(control);
        if (point.Properties.IsRightButtonPressed)
        {
            ShowMenu(control, chip.Occurrence);
            return;
        }
        if (!point.Properties.IsLeftButtonPressed || e.ClickCount > 1 || !CalendarViewModel.CanMove(chip.Occurrence))
            return;

        // The same gesture works on the month grid (7 x 6 days) and the all-day strip (one row).
        var inMonth = MonthGrid.IsVisualAncestorOf(control);
        Visual surface = inMonth ? MonthGrid : AllDayStrip;
        var days = inMonth ? vm.MonthCells.Select(c => c.Date).ToList() : vm.Days.Select(d => d.Date).ToList();
        var columns = inMonth ? 7 : Math.Max(1, vm.Days.Count);
        var rows = inMonth ? 6 : 1;
        var origin = e.GetPosition(surface);
        _drag = new DragState
        {
            Source = control,
            Pointer = e.Pointer,
            Occurrence = chip.Occurrence,
            Kind = DragKind.MoveDays,
            Surface = surface,
            Origin = origin,
            Days = days,
            Columns = columns,
            Rows = rows,
            StartIndex = CellAt(surface, origin, columns, rows),
        };
        e.Pointer.Capture(control);
    }

    private void OnChipPointerMoved(object? sender, PointerEventArgs e)
    {
        if (sender is not Control control) return;
        if (_drag is null)
        {
            control.Cursor = control.DataContext is EventChipViewModel chip && CalendarViewModel.CanMove(chip.Occurrence)
                ? MoveCursor : Cursor.Default;
            return;
        }
        if (!ReferenceEquals(_drag.Source, control)) return;

        var position = e.GetPosition(_drag.Surface);
        if (!_drag.Active)
        {
            var moved = position - _drag.Origin;
            if (Math.Abs(moved.X) < DragThreshold && Math.Abs(moved.Y) < DragThreshold) return;
            _drag.Active = true;
            control.Opacity = 0.45;
        }

        var cell = CellAt(_drag.Surface, position, _drag.Columns, _drag.Rows);
        var days = _drag.Days[cell] - _drag.Days[_drag.StartIndex];
        var o = _drag.Occurrence;
        if (o.IsAllDay)
        {
            // All-day items are floating dates: shift the date, not the instant.
            _drag.NewStart = o.Start.AddDays(days.Days);
            _drag.NewEnd = o.End.AddDays(days.Days);
        }
        else
        {
            // Keep the time of day across daylight-saving changes.
            _drag.NewStart = ToOffset(o.Start.ToLocalTime().DateTime + days);
            _drag.NewEnd = ToOffset(o.End.ToLocalTime().DateTime + days);
        }

        var ghost = ReferenceEquals(_drag.Surface, MonthGrid) ? MonthGhost : AllDayGhost;
        var width = _drag.Surface.Bounds.Width / _drag.Columns;
        var height = _drag.Surface.Bounds.Height / _drag.Rows;
        Canvas.SetLeft(ghost, cell % _drag.Columns * width + 1);
        Canvas.SetTop(ghost, cell / _drag.Columns * height + 1);
        ghost.Width = Math.Max(0, width - 2);
        ghost.Height = Math.Max(0, height - 2);
        ghost.IsVisible = true;
    }

    private async void OnChipReleased(object? sender, PointerReleasedEventArgs e) => await FinishDragAsync();

    private static int CellAt(Visual surface, Point position, int columns, int rows)
    {
        var column = Math.Clamp((int)Math.Floor(position.X / (surface.Bounds.Width / columns)), 0, columns - 1);
        var row = Math.Clamp((int)Math.Floor(position.Y / (surface.Bounds.Height / rows)), 0, rows - 1);
        return row * columns + column;
    }

    private async void OnChipDoubleTapped(object? sender, TappedEventArgs e)
    {
        e.Handled = true;
        if (sender is Control { DataContext: EventChipViewModel chip } && Vm is { } vm)
            await vm.OpenAsync(chip.Occurrence);
    }

    private void OnMonthCellDoubleTapped(object? sender, TappedEventArgs e)
    {
        if (sender is Control { DataContext: MonthCellViewModel cell } && Vm is { } vm)
            vm.GoToDay(cell.Date);
    }

    // ───────────────────────────── Ending a drag ─────────────────────────────

    /// <summary>Drop: save the new time if it changed. A press without movement was just a click.</summary>
    private async Task FinishDragAsync()
    {
        var drag = _drag;
        if (drag is null) return;
        EndDrag();
        if (drag.Active && Vm is { } vm)
            await vm.MoveAsync(drag.Occurrence, drag.NewStart, drag.NewEnd);
    }

    /// <summary>Escape, or losing the pointer (another window, a dialog): put everything back.</summary>
    private void OnDragCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        if (_drag is not null && ReferenceEquals(_drag.Source, sender))
            EndDrag();
    }

    private void OnKeyDownTunnel(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape && _drag is not null)
        {
            EndDrag();
            e.Handled = true;
        }
    }

    private void EndDrag()
    {
        var drag = _drag;
        _drag = null;   // first, so the capture-lost event below finds nothing to cancel
        if (drag is null) return;
        drag.Source.Opacity = 1;
        TimeGhost.IsVisible = false;
        MonthGhost.IsVisible = false;
        AllDayGhost.IsVisible = false;
        drag.Pointer.Capture(null);
    }

    private static DateTime Round(DateTime value) =>
        new(value.Ticks + Snap.Ticks / 2 - (value.Ticks + Snap.Ticks / 2) % Snap.Ticks, value.Kind);

    private static DateTime Min(DateTime a, DateTime b) => a < b ? a : b;
    private static DateTime Max(DateTime a, DateTime b) => a > b ? a : b;

    private static DateTimeOffset ToOffset(DateTime local) =>
        new(DateTime.SpecifyKind(local, DateTimeKind.Local));

    // ───────────────────────────── Menus and dialogs ─────────────────────────────

    private void ShowMenu(Control target, AppointmentOccurrence occurrence)
    {
        if (Vm is not { } vm) return;
        var open = new MenuItem { Header = "Open" };
        open.Click += async (_, _) => await vm.OpenAsync(occurrence);
        var delete = new MenuItem { Header = occurrence.Appointment.IsOrganizer && occurrence.Appointment.IsMeeting ? "Cancel meeting" : "Delete" };
        delete.Click += async (_, _) => await vm.DeleteAsync(occurrence);
        new ContextMenu { ItemsSource = new[] { open, delete } }.Open(target);
    }

    /// <summary>Outlook's question for repeating items. True = this occurrence, false = the series, null = cancel.</summary>
    private async Task<bool?> AskOccurrenceOrSeriesAsync(AppointmentOccurrence occurrence)
    {
        if (TopLevel.GetTopLevel(this) is not Window owner) return false;
        bool? result = null;
        var dialog = new Window
        {
            Title = "Delete repeating item",
            Width = 400,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false,
        };
        var single = new Button { Content = "This occurrence", Classes = { "primary" }, IsDefault = true };
        var series = new Button { Content = "The entire series" };
        var cancel = new Button { Content = "Cancel", IsCancel = true };
        single.Click += (_, _) => { result = true; dialog.Close(); };
        series.Click += (_, _) => { result = false; dialog.Close(); };
        cancel.Click += (_, _) => dialog.Close();
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 16,
            Children =
            {
                new TextBlock
                {
                    Text = $"“{occurrence.Appointment.Title}” repeats. Delete only the occurrence on {occurrence.Start.ToLocalTime():D}, or every occurrence?",
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                },
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8, Children = { single, series, cancel } },
            },
        };
        await dialog.ShowDialog(owner);
        return result;
    }
}
