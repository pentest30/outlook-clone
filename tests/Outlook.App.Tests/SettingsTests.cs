using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;
using Outlook.App.ViewModels;
using Outlook.App.Views;

namespace Outlook.App.Tests;

public class SettingsTests
{
    private static AppSettings Fresh(out string path)
    {
        path = Path.Combine(Path.GetTempPath(), $"settings-{Guid.NewGuid():N}.json");
        return AppSettings.Load(path);
    }

    [AvaloniaFact]
    public void Save_applies_persists_and_notifies()
    {
        var settings = Fresh(out var path);
        var changed = 0;
        settings.Changed += () => changed++;
        var vm = new SettingsViewModel(settings, "me@x.test")
        {
            Theme = "Dark",
            Signature = "Iqbal Baouche\nNetcom Group\n\n",
            UndoSendSeconds = 8,
            MarkReadWhenOpened = false,
            DefaultReminder = SettingsViewModel.ReminderOptions.First(r => r.Minutes == 30),
        };

        vm.SaveCommand.Execute(null);

        Assert.Equal(1, changed);
        var reloaded = AppSettings.Load(path);
        Assert.Equal("Dark", reloaded.Theme);
        Assert.Equal("Iqbal Baouche\nNetcom Group", reloaded.SignatureFor("ME@x.test"));   // trimmed; address case ignored
        Assert.Equal(8, reloaded.UndoSendSeconds);
        Assert.False(reloaded.MarkReadWhenOpened);
        Assert.Equal(30, reloaded.DefaultReminderMinutes);
    }

    [AvaloniaFact]
    public void Cancel_changes_nothing()
    {
        var settings = Fresh(out _);
        var vm = new SettingsViewModel(settings, "me@x.test") { Theme = "Dark", Signature = "x" };
        vm.CancelCommand.Execute(null);
        Assert.Equal("System", settings.Theme);
        Assert.Equal("", settings.SignatureFor("me@x.test"));
    }

    [AvaloniaFact]
    public void Clearing_the_signature_removes_it()
    {
        var settings = Fresh(out _);
        settings.Signatures["me@x.test"] = "old";
        new SettingsViewModel(settings, "me@x.test") { Signature = "  " }.SaveCommand.Execute(null);
        Assert.False(settings.Signatures.ContainsKey("me@x.test"));
    }

    [AvaloniaFact]
    public void Window_renders_every_tab()
    {
        var settings = Fresh(out _);
        settings.Signatures["me@x.test"] = "Iqbal Baouche\nNetcom Group";
        var window = new SettingsWindow { DataContext = new SettingsViewModel(settings, "me@x.test") };
        window.Show();
        var tabs = window.FindControl<TabControl>("Tabs")!;
        var shots = Path.Combine(Path.GetTempPath(), "outlook-ui-shots");
        Directory.CreateDirectory(shots);
        for (var i = 0; i < tabs.ItemCount; i++)
        {
            tabs.SelectedIndex = i;
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Assert.NotNull(window.CaptureRenderedFrame());
            window.CaptureRenderedFrame()!.Save(Path.Combine(shots, $"settings-{i}.png"));
        }
    }
}
