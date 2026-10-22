using CommunityToolkit.Mvvm.ComponentModel;
using Outlook.Core.Models;

namespace Outlook.App.ViewModels;

public partial class FolderViewModel(MailFolder folder, int depth) : ObservableObject
{
    public MailFolder Model { get; } = folder;
    public int Id => Model.Id;
    public string Name => Model.DisplayName;

    /// <summary>Indent for nested folders in the folder pane.</summary>
    public Avalonia.Thickness Indent { get; } = new(12 + depth * 16, 0, 8, 0);

    /// <summary>The folder's icon (Styles/Icons.axaml), as in Outlook's folder pane.</summary>
    public Avalonia.Media.Geometry? Icon => Icons.Get(Model.Special switch
    {
        SpecialFolder.Inbox => "MailInbox",
        SpecialFolder.Drafts => "Drafts",
        SpecialFolder.Sent => "Send",
        SpecialFolder.Trash => "Delete",
        SpecialFolder.Junk => "Warning",
        SpecialFolder.Archive => "Archive",
        _ => "Folder",
    });

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasUnread))]
    private int _unreadCount = folder.UnreadCount;

    public bool HasUnread => UnreadCount > 0;
}
