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

    public string Icon => Model.Special switch
    {
        SpecialFolder.Inbox => "📥",
        SpecialFolder.Drafts => "📝",
        SpecialFolder.Sent => "📤",
        SpecialFolder.Trash => "🗑",
        SpecialFolder.Junk => "⚠",
        SpecialFolder.Archive => "🗄",
        _ => "📁",
    };

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasUnread))]
    private int _unreadCount = folder.UnreadCount;

    public bool HasUnread => UnreadCount > 0;
}
