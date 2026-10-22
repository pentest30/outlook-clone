namespace Outlook.App.Tests;

public class SingleInstanceTests
{
    [Fact]
    public void A_second_launch_asks_the_first_to_show_itself_and_leaves()
    {
        var dataDir = Path.Combine(Path.GetTempPath(), $"outlook-instance-{Guid.NewGuid():N}");
        using var first = SingleInstance.TryStart(dataDir);
        Assert.NotNull(first);
        using var shown = new ManualResetEventSlim();
        first.ShowRequested += shown.Set;

        Assert.Null(SingleInstance.TryStart(dataDir));
        Assert.True(shown.Wait(TimeSpan.FromSeconds(5)));

        // Another data folder (OUTLOOKCLONE_DATA) is another instance.
        using var other = SingleInstance.TryStart(dataDir + "-other");
        Assert.NotNull(other);
    }
}
