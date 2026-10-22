using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

/// <summary>What the reading pane and the message window do with attachments and links in a mail body.</summary>
public static class MailActions
{
    /// <summary>Opening these from mail is how malware spreads; Outlook blocks them too. Saving is still allowed.</summary>
    private static readonly HashSet<string> BlockedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".exe", ".com", ".bat", ".cmd", ".msi", ".msp", ".scr", ".pif", ".cpl", ".js", ".jse", ".vbs", ".vbe",
        ".wsf", ".wsh", ".ps1", ".psm1", ".hta", ".jar", ".lnk", ".reg", ".inf", ".application", ".appref-ms",
    };

    public static bool CanOpen(string fileName) => !BlockedExtensions.Contains(Path.GetExtension(fileName));

    /// <summary>Open or Save as, under the attachment chip.</summary>
    public static void ShowAttachmentMenu(Control anchor, AttachmentViewModel attachment)
    {
        var canOpen = CanOpen(attachment.FileName);
        var open = new MenuItem { Header = canOpen ? "Open" : "Open (blocked: this file type can run code)", IsEnabled = canOpen };
        open.Click += async (_, _) => await OpenAttachmentAsync(anchor, attachment);
        var save = new MenuItem { Header = "Save as…" };
        save.Click += async (_, _) => await SaveAttachmentAsync(anchor, attachment);
        new ContextMenu { ItemsSource = new[] { open, save }, Placement = PlacementMode.BottomEdgeAlignedLeft }.Open(anchor);
    }

    private static async Task OpenAttachmentAsync(Control anchor, AttachmentViewModel attachment)
    {
        if (TopLevel.GetTopLevel(anchor) is not { } top || attachment.LoadData is not { } load) return;
        // A fresh folder per open, so two attachments with the same name never collide.
        var dir = Path.Combine(Path.GetTempPath(), "OutlookClone", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, attachment.FileName);
        await File.WriteAllBytesAsync(path, await load());
        await top.Launcher.LaunchFileInfoAsync(new FileInfo(path));
    }

    private static async Task SaveAttachmentAsync(Control anchor, AttachmentViewModel attachment)
    {
        if (TopLevel.GetTopLevel(anchor) is not { } top || attachment.LoadData is not { } load) return;
        var file = await top.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save attachment",
            SuggestedFileName = attachment.FileName,
            ShowOverwritePrompt = true,
        });
        if (file is null) return;
        await using var stream = await file.OpenWriteAsync();
        await stream.WriteAsync(await load());
    }

    /// <summary>mailto: starts a message; web links open in the browser; anything else is ignored.</summary>
    public static async Task FollowLinkAsync(TopLevel top, string link, Action<string> composeTo)
    {
        if (link.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
            composeTo(Uri.UnescapeDataString(link[7..].Split('?')[0]));
        else if (Uri.TryCreate(link, UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https")
            await top.Launcher.LaunchUriAsync(uri);
    }
}
