using System.Text.Json;

namespace Outlook.App;

/// <summary>User preferences, kept in settings.json next to the cache.</summary>
public sealed class AppSettings
{
    /// <summary>Outlook's "Show as Conversations".</summary>
    public bool ShowAsConversations { get; set; } = true;

    /// <summary>How long a sent message waits with an Undo button. 0 sends at once. Outlook offers up to 10.</summary>
    public int UndoSendSeconds { get; set; } = 5;

    /// <summary>Whether search covers every folder (Outlook's "All mailboxes") or just the open one.</summary>
    public bool SearchAllFolders { get; set; } = true;

    private string? _path;

    public static AppSettings Load(string path)
    {
        AppSettings settings;
        try
        {
            settings = File.Exists(path) ? JsonSerializer.Deserialize<AppSettings>(File.ReadAllText(path)) ?? new() : new();
        }
        catch (Exception ex) when (ex is JsonException or IOException)
        {
            settings = new();   // a damaged file should not stop the app from starting
        }
        settings._path = path;
        settings.UndoSendSeconds = Math.Clamp(settings.UndoSendSeconds, 0, 10);
        return settings;
    }

    public void Save()
    {
        if (_path is null) return;
        try
        {
            File.WriteAllText(_path, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (IOException)
        {
            // Preferences are a convenience; never fail an action over them.
        }
    }
}
