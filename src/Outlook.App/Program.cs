using Avalonia;
using Velopack;

namespace Outlook.App;

internal static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // First thing: lets the installer run its install / update / uninstall hooks (shortcuts),
        // which start the app with special arguments and exit straight away.
        VelopackApp.Build().Run();
        // Already running (perhaps hidden in the notification area): it comes forward, we leave.
        using var instance = SingleInstance.TryStart(App.DataDir);
        if (instance is null) return;
        App.Instance = instance;
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
