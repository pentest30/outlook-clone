using Avalonia;
using Avalonia.Media;

namespace Outlook.App;

/// <summary>Looks up an icon shape from Styles/Icons.axaml by name, for icons chosen in code.</summary>
public static class Icons
{
    public static Geometry? Get(string name) =>
        Application.Current?.TryGetResource("Icon." + name, null, out var value) == true ? value as Geometry : null;
}
