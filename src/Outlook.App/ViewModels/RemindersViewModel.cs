using System.Collections.ObjectModel;
using System.Globalization;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Models;
using Outlook.Core.Services;

namespace Outlook.App.ViewModels;

public sealed partial class ReminderItemViewModel(AppointmentOccurrence occurrence) : ObservableObject
{
    public AppointmentOccurrence Occurrence { get; } = occurrence;
    public string Key { get; } = $"{occurrence.Appointment.Id}@{occurrence.Start.UtcTicks}";
    public string Title => Occurrence.Appointment.Title;
    public string When => Occurrence.IsAllDay
        ? Occurrence.Start.UtcDateTime.ToString("D", CultureInfo.CurrentCulture)
        : Occurrence.Start.ToLocalTime().ToString("f", CultureInfo.CurrentCulture);
    public string Location => Occurrence.Appointment.Location;

    [ObservableProperty] private string _dueText = "";

    public void Refresh(DateTimeOffset now)
    {
        var until = Occurrence.Start - now;
        DueText = until <= TimeSpan.Zero ? "Now"
            : until < TimeSpan.FromHours(1) ? $"In {Math.Ceiling(until.TotalMinutes)} minutes"
            : until < TimeSpan.FromDays(1) ? $"In {Math.Round(until.TotalHours)} hours"
            : $"In {Math.Round(until.TotalDays)} days";
    }
}

/// <summary>
/// Outlook's Reminders window: checks every 30 seconds for appointments whose reminder time has
/// come, lists them until dismissed, snoozes for 5 minutes.
/// </summary>
public partial class RemindersViewModel : ObservableObject
{
    private static readonly TimeSpan SnoozeTime = TimeSpan.FromMinutes(5);
    /// <summary>Longest reminder offered by the editor (1 week), so the look-ahead window covers it.</summary>
    private static readonly TimeSpan LookAhead = TimeSpan.FromDays(8);

    private readonly ICalendarService _calendar;
    private readonly Func<Account?> _account;
    private readonly DispatcherTimer _timer;
    private readonly Dictionary<string, DateTimeOffset> _snoozedUntil = [];

    public RemindersViewModel(ICalendarService calendar, Func<Account?> account)
    {
        _calendar = calendar;
        _account = account;
        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _timer.Tick += async (_, _) => await CheckAsync();
    }

    /// <summary>Raised when the window should appear (new reminders came due).</summary>
    public event Action? ShowRequested;
    /// <summary>Raised to open an item from the window.</summary>
    public event Action<AppointmentOccurrence>? OpenRequested;

    public ObservableCollection<ReminderItemViewModel> Items { get; } = [];

    public void Start()
    {
        _timer.Start();
        _ = CheckAsync();
    }

    public async Task CheckAsync()
    {
        if (_account() is not { } account) return;
        var now = DateTimeOffset.Now;
        IReadOnlyList<AppointmentOccurrence> upcoming;
        try
        {
            upcoming = await _calendar.GetOccurrencesAsync(account.Id, now.AddHours(-1), now + LookAhead);
        }
        catch (Exception)
        {
            return;   // try again on the next tick
        }

        var added = false;
        foreach (var o in upcoming)
        {
            if (o.Appointment.ReminderMinutes is not { } minutes) continue;
            if (o.End <= now) continue;                                            // already over
            var start = o.IsAllDay ? new DateTimeOffset(DateTime.SpecifyKind(o.Start.UtcDateTime, DateTimeKind.Local)) : o.Start;
            if (start - TimeSpan.FromMinutes(minutes) > now) continue;             // not yet
            if (o.Appointment.RemindersDismissedThroughUtc >= o.Start) continue;  // dismissed
            var item = new ReminderItemViewModel(o);
            if (_snoozedUntil.TryGetValue(item.Key, out var until) && until > now) continue;
            if (Items.Any(i => i.Key == item.Key)) continue;
            Items.Add(item);
            added = true;
        }
        foreach (var item in Items) item.Refresh(now);
        if (added) ShowRequested?.Invoke();
    }

    [RelayCommand]
    private async Task DismissAsync(ReminderItemViewModel item)
    {
        Items.Remove(item);
        await _calendar.DismissReminderAsync(item.Occurrence.Appointment, item.Occurrence.Start);
    }

    [RelayCommand]
    private async Task DismissAllAsync()
    {
        foreach (var item in Items.ToList())
            await DismissAsync(item);
    }

    [RelayCommand]
    private void Snooze(ReminderItemViewModel item)
    {
        Items.Remove(item);
        _snoozedUntil[item.Key] = DateTimeOffset.Now + SnoozeTime;
    }

    [RelayCommand]
    private void Open(ReminderItemViewModel item) => OpenRequested?.Invoke(item.Occurrence);
}
