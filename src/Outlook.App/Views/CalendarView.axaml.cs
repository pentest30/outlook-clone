using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Outlook.App.ViewModels;
using Outlook.Core.Models;

namespace Outlook.App.Views;

public partial class CalendarView : UserControl
{
    public CalendarView()
    {
        InitializeComponent();
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

    // ───────────────────────────── Time grid ─────────────────────────────

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

    private void OnBlockPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control { DataContext: EventBlockViewModel block } control || Vm is not { } vm) return;
        vm.Selected = block.Occurrence;
        Focus();   // so Delete / Enter reach the calendar
        if (e.GetCurrentPoint(control).Properties.IsRightButtonPressed)
            ShowMenu(control, block.Occurrence);
        e.Handled = true;
    }

    private async void OnBlockDoubleTapped(object? sender, TappedEventArgs e)
    {
        e.Handled = true;   // not a double-click on the empty column underneath
        if (sender is Control { DataContext: EventBlockViewModel block } && Vm is { } vm)
            await vm.OpenAsync(block.Occurrence);
    }

    // ───────────────────────────── All-day chips and month view ─────────────────────────────

    private void OnChipPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control { DataContext: EventChipViewModel chip } control || Vm is not { } vm) return;
        vm.Selected = chip.Occurrence;
        Focus();
        if (e.GetCurrentPoint(control).Properties.IsRightButtonPressed)
            ShowMenu(control, chip.Occurrence);
        e.Handled = true;
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
