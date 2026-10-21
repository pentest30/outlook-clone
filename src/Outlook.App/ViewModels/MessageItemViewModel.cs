using System.Globalization;
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using Outlook.Core.Models;

namespace Outlook.App.ViewModels;

public record CategoryChip(string Name, IBrush Brush, string Keyword);

/// <summary>
/// One row of the message list: a message, or in conversation view a whole thread (represented by
/// its newest message in this folder).
/// </summary>
public partial class MessageItemViewModel : ObservableObject
{
    public MessageItemViewModel(MailMessage model)
    {
        _model = model;
        ThreadInFolder = [model];
        ThreadAll = [model];
        _isRead = model.IsRead;
        _isFlagged = model.IsFlagged;
    }

    [ObservableProperty]
    private MailMessage _model;

    /// <summary>Takes fresh data from the cache while keeping this item (and so the list selection).</summary>
    public void Update(MailMessage fresh)
    {
        Model = fresh;
        ThreadInFolder = [fresh];
        ThreadAll = [fresh];
        IsRead = fresh.IsRead;
        IsFlagged = fresh.IsFlagged;
        OnPropertyChanged(string.Empty);
    }

    /// <summary>Conversation view: the thread's messages in this folder and everywhere, newest first.</summary>
    public void SetThread(IReadOnlyList<MailMessage> inFolder, IReadOnlyList<MailMessage> all)
    {
        ThreadInFolder = inFolder;
        ThreadAll = all;
        Model = inFolder[0];
        IsRead = inFolder.All(m => m.IsRead);
        IsFlagged = inFolder.Any(m => m.IsFlagged);
        OnPropertyChanged(string.Empty);
    }

    /// <summary>The thread key used to match rows across refreshes in conversation view.</summary>
    public string Key => IsConversationRow ? "t:" + Model.ThreadId : "m:" + Model.Id;
    public bool IsConversationRow { get; set; }

    public IReadOnlyList<MailMessage> ThreadInFolder { get; private set; }
    public IReadOnlyList<MailMessage> ThreadAll { get; private set; }
    public int ThreadCount => ThreadAll.Count;
    public bool ShowThreadCount => ThreadCount > 1;

    /// <summary>"Ana Lopez; Bilal Haddad" for a conversation, the sender for a single message.</summary>
    public string Sender => ThreadAll.Count > 1
        ? string.Join("; ", ThreadAll.Select(m => m.SenderDisplay).Distinct(StringComparer.OrdinalIgnoreCase).Take(4))
        : Model.SenderDisplay;

    public string Subject => string.IsNullOrWhiteSpace(Model.Subject) ? "(no subject)" : Model.Subject;
    public string Preview => Model.Preview;
    public bool HasAttachments => ThreadInFolder.Any(m => m.HasAttachments);
    public bool IsAnswered => Model.IsAnswered;
    public bool IsForwarded => Model.IsForwarded;
    public string To => Model.To;
    public string Cc => Model.Cc;
    public bool HasCc => !string.IsNullOrWhiteSpace(Model.Cc);
    public string FromLine => string.IsNullOrWhiteSpace(Model.FromName) ? Model.FromAddress : $"{Model.FromName} <{Model.FromAddress}>";
    public string Initials => InitialsFor(Model.SenderDisplay);

    /// <summary>Search results span folders, so each row says where the message is.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ShowFolderName))]
    private string? _folderName;

    public bool ShowFolderName => FolderName is not null;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasCategories))]
    private IReadOnlyList<CategoryChip> _categories = [];

    public bool HasCategories => Categories.Count > 0;

    /// <summary>Two letters for the round avatar, e.g. "Ana Lopez" becomes "AL".</summary>
    public static string InitialsFor(string name)
    {
        var parts = name.Split([' ', '.', '_', '@'], StringSplitOptions.RemoveEmptyEntries);
        return parts.Length switch
        {
            0 => "?",
            1 => parts[0][..1].ToUpperInvariant(),
            _ => (parts[0][..1] + parts[1][..1]).ToUpperInvariant(),
        };
    }

    public DateTimeOffset LocalDate => Model.Date.ToLocalTime();

    /// <summary>Outlook-style list date: time for today, weekday this week, short date otherwise.</summary>
    public string ShortDate
    {
        get
        {
            var d = LocalDate;
            var today = DateTime.Today;
            if (d.Date == today) return d.ToString("t", CultureInfo.CurrentCulture);
            // Separate calls: inside a custom format string "t" means the AM/PM designator, not short time.
            if (d.Date > today.AddDays(-7))
                return $"{d.ToString("ddd", CultureInfo.CurrentCulture)} {d.ToString("t", CultureInfo.CurrentCulture)}";
            return d.ToString("d", CultureInfo.CurrentCulture);
        }
    }

    public string FullDate => LocalDate.ToString("f", CultureInfo.CurrentCulture);

    /// <summary>Date group label ("Today", "Yesterday", ...) and whether this item opens a new group.</summary>
    [ObservableProperty] private string _groupLabel = "";
    [ObservableProperty] private bool _showGroupHeader;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TitleWeight), nameof(ShowUnreadBar))]
    private bool _isRead;

    [ObservableProperty]
    private bool _isFlagged;

    public FontWeight TitleWeight => IsRead ? FontWeight.Normal : FontWeight.SemiBold;
    public bool ShowUnreadBar => !IsRead;

    public static string GroupFor(DateTimeOffset localDate)
    {
        var today = DateTime.Today;
        var d = localDate.Date;
        if (d >= today) return "Today";
        if (d == today.AddDays(-1)) return "Yesterday";
        var startOfWeek = today.AddDays(-(((int)today.DayOfWeek - (int)CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek + 7) % 7));
        if (d >= startOfWeek) return "This week";
        if (d >= startOfWeek.AddDays(-7)) return "Last week";
        if (d >= new DateTime(today.Year, today.Month, 1)) return "Earlier this month";
        if (d >= new DateTime(today.Year, today.Month, 1).AddMonths(-1)) return "Last month";
        return "Older";
    }
}
