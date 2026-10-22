using System.Text.Json;

namespace Outlook.App;

/// <summary>User preferences, kept in settings.json next to the cache. Edited in the Settings window.</summary>
public sealed class AppSettings
{
    // ── General ──
    /// <summary>"System" (follow Windows), "Light" or "Dark".</summary>
    public string Theme { get; set; } = "System";

    /// <summary>Desktop alerts for new mail while the window is in the background.</summary>
    public bool NewMailNotifications { get; set; } = true;

    /// <summary>The account open in the window, remembered between runs.</summary>
    public int? ActiveAccountId { get; set; }

    // ── Mail ──
    /// <summary>Outlook's "Show as Conversations".</summary>
    public bool ShowAsConversations { get; set; } = true;

    /// <summary>Outlook's "Mark item as read when selection changes"; off means only Ctrl+Q marks read.</summary>
    public bool MarkReadWhenOpened { get; set; } = true;

    /// <summary>How long a sent message waits with an Undo button. 0 sends at once. Outlook offers up to 10.</summary>
    public int UndoSendSeconds { get; set; } = 5;

    /// <summary>Whether search covers every folder (Outlook's "All mailboxes") or just the open one.</summary>
    public bool SearchAllFolders { get; set; } = true;

    // ── Signatures ──
    /// <summary>Signature text per account e-mail address.</summary>
    public Dictionary<string, string> Signatures { get; set; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Also add the signature to replies and forwards (always added to new messages).</summary>
    public bool SignatureOnReplies { get; set; } = true;

    // ── Calendar ──
    /// <summary>Reminder for new appointments, in minutes; null for none.</summary>
    public int? DefaultReminderMinutes { get; set; } = 15;

    /// <summary>Raised after the Settings window saves, so open views can apply the changes.</summary>
    public event Action? Changed;

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
        // JSON loses the case-insensitive comparer.
        settings.Signatures = new Dictionary<string, string>(settings.Signatures, StringComparer.OrdinalIgnoreCase);
        return settings;
    }

    public string SignatureFor(string email) => Signatures.GetValueOrDefault(email) ?? "";

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

    /// <summary>Saves and tells listeners (theme, list view, compose defaults) to pick up the new values.</summary>
    public void SaveAndNotify()
    {
        Save();
        Changed?.Invoke();
    }
}
