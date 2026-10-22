using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Outlook.App.ViewModels;
using Outlook.App.Views;
using Outlook.Core.Compose;
using Outlook.Core.Models;

namespace Outlook.App.Tests;

public class AutocompleteTests
{
    private static readonly string[] People = ["Ana Lopez <ana@contoso.test>", "Andrew Kim <andrew@fabrikam.test>", "Bilal Haddad <bilal@fabrikam.test>"];

    /// <summary>Same contract as the real suggester: match the entry being typed.</summary>
    private static Task<IEnumerable<object>> Suggest(string? line, CancellationToken ct)
    {
        var typed = AddressList.LastEntry(line).ToLowerInvariant();
        return Task.FromResult<IEnumerable<object>>(typed.Length == 0 ? [] : People.Where(p => p.ToLowerInvariant().Contains(typed)).ToList());
    }

    /// <summary>Runs the UI loop until a condition holds (the suggestion list waits a short delay before asking).</summary>
    private static void WaitFor(Func<bool> condition)
    {
        var until = DateTime.UtcNow.AddSeconds(3);
        while (!condition() && DateTime.UtcNow < until)
        {
            Dispatcher.UIThread.RunJobs();
            AvaloniaHeadlessPlatform.ForceRenderTimerTick();
            Thread.Sleep(20);
        }
    }

    [AvaloniaFact]
    public void Typing_suggests_people_and_picking_one_completes_the_entry()
    {
        var account = new Account { Id = 1, Email = "me@x.test", DisplayName = "Me" };
        var compose = new ComposeViewModel(null!, account, ComposeMode.New, null,
            new ComposeSeed("bob@y.test; ", "", "", "", null, null), [], null) { SuggestRecipients = Suggest };
        var window = new ComposeWindow { DataContext = compose };
        window.Show();
        var to = window.FindControl<AutoCompleteBox>("ToBox")!;
        // The real window waits 150 ms between keystrokes before looking people up; that wait runs on a
        // dispatcher timer, which the headless test loop does not advance, so look up straight away here.
        to.MinimumPopulateDelay = TimeSpan.Zero;
        Dispatcher.UIThread.RunJobs();
        // Type into the text field inside the box, at the end of what is already there (as after a click).
        var field = to.GetVisualDescendants().OfType<TextBox>().First();
        field.Focus();
        field.CaretIndex = field.Text!.Length;
        Dispatcher.UIThread.RunJobs();

        window.KeyTextInput("an");
        WaitFor(() => to.IsDropDownOpen);

        Assert.True(to.IsDropDownOpen,
            $"text='{to.Text}' search='{to.SearchText}' items={to.ItemsSource?.Cast<object>().Count()} focused={to.IsKeyboardFocusWithin} vmTo='{compose.To}' populator={to.AsyncPopulator is not null}");
        var shown = to.ItemsSource!.Cast<object>().Select(o => o.ToString()).ToList();
        Assert.Equal(["Ana Lopez <ana@contoso.test>", "Andrew Kim <andrew@fabrikam.test>"], shown);
        Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "outlook-ui-shots"));
        window.CaptureRenderedFrame()?.Save(Path.Combine(Path.GetTempPath(), "outlook-ui-shots", "autocomplete.png"));

        // Pick the first suggestion with the keyboard.
        window.KeyPressQwerty(PhysicalKey.ArrowDown, RawInputModifiers.None);
        window.KeyPressQwerty(PhysicalKey.Enter, RawInputModifiers.None);
        WaitFor(() => compose.To.Contains("Ana Lopez"));

        Assert.Equal("bob@y.test; Ana Lopez <ana@contoso.test>; ", compose.To);
        compose.DiscardCommand.Execute(null);   // stop the auto-save timer
        window.Close();
    }
}
