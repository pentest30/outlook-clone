using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Models;

namespace Outlook.App.Tests;

public class NotificationTests
{
    private static MailMessage Mail(int id, string from, string subject) =>
        new() { Id = id, FromName = from, FromAddress = $"{from.ToLowerInvariant()}@x.test", Subject = subject, Preview = $"Preview of {subject}" };

    [AvaloniaFact]
    public void Shows_at_most_three_alerts_and_summarises_the_rest()
    {
        var owner = new Window();
        owner.Show();
        var toasts = new ToastManager(owner, _ => { });

        toasts.ShowNewMail([Mail(1, "Ana", "Budget"), Mail(2, "Bilal", "Invoice"), Mail(3, "Chen", "Q3"), Mail(4, "Dana", "Hi"), Mail(5, "Eve", "Yo")]);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(3, toasts.Open.Count);
        var contents = toasts.Open.Select(t => (ToastContent)t.DataContext!).ToList();
        Assert.Equal(["Ana", "Bilal", "Chen"], contents.Select(c => c.Sender));
        Assert.Equal("New message · +2 more new messages", contents[^1].Heading);
        Assert.Equal("New message", contents[0].Heading);

        AvaloniaHeadlessPlatform.ForceRenderTimerTick();
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "outlook-ui-shots"));
        toasts.Open[0].CaptureRenderedFrame()?.Save(Path.Combine(Path.GetTempPath(), "outlook-ui-shots", "toast.png"));
    }

    [AvaloniaFact]
    public void Clicking_an_alert_opens_its_message_and_closes_it()
    {
        var owner = new Window();
        owner.Show();
        MailMessage? opened = null;
        var toasts = new ToastManager(owner, m => opened = m);
        toasts.ShowNewMail([Mail(7, "Ana", "Budget")]);
        Dispatcher.UIThread.RunJobs();

        toasts.Open[0].Click();
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(7, opened?.Id);
        Assert.Empty(toasts.Open);
    }

    [AvaloniaFact]
    public void New_alerts_push_out_the_oldest()
    {
        var owner = new Window();
        owner.Show();
        var toasts = new ToastManager(owner, _ => { });
        toasts.ShowNewMail([Mail(1, "Ana", "a"), Mail(2, "Bilal", "b"), Mail(3, "Chen", "c")]);
        toasts.ShowNewMail([Mail(4, "Dana", "d")]);
        Dispatcher.UIThread.RunJobs();

        Assert.Equal(["Bilal", "Chen", "Dana"], toasts.Open.Select(t => ((ToastContent)t.DataContext!).Sender));
    }

    [AvaloniaFact]
    public void A_due_reminder_shows_an_alert_that_opens_it()
    {
        var owner = new Window();
        owner.Show();
        var toasts = new ToastManager(owner, _ => { });
        var start = DateTimeOffset.Now.AddMinutes(10);
        var appointment = new Appointment { Id = 3, Title = "Equipe SI Billing", Location = "Webex", StartUtc = start, EndUtc = start.AddHours(1) };
        var reminder = new ReminderItemViewModel(new AppointmentOccurrence(appointment, start, start.AddHours(1)));
        reminder.Refresh(DateTimeOffset.Now);
        ReminderItemViewModel? opened = null;

        toasts.ShowReminders([reminder], r => opened = r);
        Dispatcher.UIThread.RunJobs();

        var content = (ToastContent)toasts.Open.Single().DataContext!;
        Assert.Equal("Reminder · In 10 minutes", content.Heading);
        Assert.Equal("Equipe SI Billing", content.Sender);
        Assert.Equal("Webex", content.Preview);

        toasts.Open[0].Click();
        Assert.Same(reminder, opened);
        Assert.Empty(toasts.Open);
    }
}
