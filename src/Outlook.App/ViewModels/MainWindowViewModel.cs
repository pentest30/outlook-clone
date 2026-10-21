using System.Collections.ObjectModel;
using Avalonia.Media;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Outlook.Core.Compose;
using Outlook.Core.Models;
using Outlook.Core.Search;
using Outlook.Core.Services;
using Outlook.Core.Text;

namespace Outlook.App.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private static readonly TimeSpan SyncInterval = TimeSpan.FromMinutes(2);

    /// <summary>Colours handed to new categories in turn, after Outlook's palette.</summary>
    private static readonly string[] NewCategoryColors = ["#00B7C3", "#C239B3", "#498205", "#DA3B01", "#8E562E", "#5C2E91"];

    private readonly IMailService _mail;
    private readonly IMailStore _store;
    private readonly AppSettings _settings;
    private readonly DispatcherTimer _syncTimer;
    private Account? _account;
    private CancellationTokenSource? _pushCts;
    private string? _loadedListKey;   // what the list currently shows: a folder (and view) or a search
    private List<MessageItemViewModel> _selection = [];
    private bool _rebuildingFolders;
    private int _reloadGeneration;
    private MailSearchQuery? _activeSearch;
    private Dictionary<string, Category> _categoriesByKeyword = [];

    public MainWindowViewModel(IMailService mail, IMailStore store, ICredentialStore credentials, AppSettings settings)
    {
        _mail = mail;
        _store = store;
        _settings = settings;
        _showAsConversations = settings.ShowAsConversations;
        _searchAllFolders = settings.SearchAllFolders;
        AccountSetup = new AccountSetupViewModel(mail, store, credentials, OnAccountAddedAsync);
        ReadingBody = new MessageBodyViewModel(mail, store);
        _syncTimer = new DispatcherTimer { Interval = SyncInterval };
        _syncTimer.Tick += (_, _) => SyncCommand.Execute(null);
        _mail.FolderChanged += (_, e) => Dispatcher.UIThread.Post(() => _ = OnServerFolderChangedAsync(e));
        _ = InitializeAsync();
    }

    /// <summary>Raised to open a compose window; the view owns windows.</summary>
    public event Action<ComposeViewModel>? ComposeRequested;

    /// <summary>Set by the view: shows a yes/no question and returns the answer.</summary>
    public Func<string, Task<bool>>? Confirm { get; set; }

    public AccountSetupViewModel AccountSetup { get; }
    public ObservableCollection<FolderViewModel> Folders { get; } = [];
    public ObservableCollection<MessageItemViewModel> Messages { get; } = [];
    public ObservableCollection<ConversationCardViewModel> ConversationCards { get; } = [];
    public ObservableCollection<PendingSendViewModel> PendingSends { get; } = [];
    public IReadOnlyList<Category> Categories { get; private set; } = [];

    /// <summary>The single message shown in the reading pane (when not showing a conversation).</summary>
    public MessageBodyViewModel ReadingBody { get; }

    [ObservableProperty] private bool _isAccountSetupVisible;
    [ObservableProperty] private string _accountEmail = "";
    [ObservableProperty] private string _statusText = "";
    [ObservableProperty] private string _folderTitle = "";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(SyncCommand))]
    private bool _isSyncing;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(NewMailCommand))]
    private FolderViewModel? _selectedFolder;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelection))]
    [NotifyCanExecuteChangedFor(nameof(MarkReadCommand), nameof(MarkUnreadCommand), nameof(ToggleReadCommand),
        nameof(ToggleFlagCommand), nameof(ReplyCommand), nameof(ReplyAllCommand), nameof(ForwardCommand),
        nameof(DeleteCommand), nameof(ArchiveCommand), nameof(MoveToCommand), nameof(OpenDraftCommand),
        nameof(ToggleCategoryCommand), nameof(ClearCategoriesCommand))]
    private MessageItemViewModel? _selectedMessage;

    /// <summary>True when the reading pane shows a conversation stack instead of one message.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSingleMessageOpen))]
    private bool _isConversationOpen;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSearchActive))]
    private string _searchText = "";

    [ObservableProperty] private bool _showAsConversations;
    [ObservableProperty] private bool _searchAllFolders;

    public bool HasSelection => SelectedMessage is not null;
    public bool IsSingleMessageOpen => HasSelection && !IsConversationOpen;
    public bool IsDraftsFolder => SelectedFolder?.Model.Special == SpecialFolder.Drafts;
    public bool IsSearchActive => _activeSearch is not null;

    // ───────────────────────────── Startup and sync ─────────────────────────────

    private async Task InitializeAsync()
    {
        var accounts = await _store.GetAccountsAsync();
        if (accounts.Count == 0)
        {
            IsAccountSetupVisible = true;
            return;
        }
        await OpenAccountAsync(accounts[0]);
    }

    private async Task OnAccountAddedAsync(Account account)
    {
        IsAccountSetupVisible = false;
        await OpenAccountAsync(account);
    }

    private async Task OpenAccountAsync(Account account)
    {
        _account = account;
        AccountEmail = account.Email;
        NewMailCommand.NotifyCanExecuteChanged();
        await LoadCategoriesAsync();
        await ReloadFoldersAsync();   // show the cache immediately, then refresh from the server
        _syncTimer.Start();
        await SyncAsync();

        _pushCts = new CancellationTokenSource();
        var token = _pushCts.Token;
        _ = Task.Run(() => _mail.RunPushAsync(account, token));
    }

    private bool CanSync() => !IsSyncing && _account is not null;

    [RelayCommand(CanExecute = nameof(CanSync))]
    private async Task SyncAsync()
    {
        if (_account is null) return;
        IsSyncing = true;
        try
        {
            var progress = new Progress<SyncProgress>(p => StatusText = $"Updating {p.FolderName}… ({p.FolderIndex + 1}/{p.FolderCount})");
            var account = _account;
            // Off the UI thread: a first sync can parse thousands of headers. Progress still reports back here.
            await Task.Run(() => _mail.SyncAccountAsync(account, progress));
            StatusText = $"All folders are up to date. Last updated {DateTime.Now:t}";
            await ReloadFoldersAsync();
        }
        catch (Exception ex)
        {
            StatusText = $"Sync failed: {ex.Message}";
        }
        finally
        {
            IsSyncing = false;
        }
    }

    /// <summary>IDLE reported new or changed mail: refresh just that folder.</summary>
    private async Task OnServerFolderChangedAsync(FolderChangedEventArgs e)
    {
        if (_account is null || e.AccountId != _account.Id) return;
        var folder = Folders.FirstOrDefault(f => f.Model.FullName == e.FolderFullName)?.Model;
        if (folder is null) return;
        try
        {
            var account = _account;
            await Task.Run(() => _mail.SyncFolderAsync(account, folder));
            await ReloadFoldersAsync();
        }
        catch (Exception ex)
        {
            StatusText = $"Couldn't update {folder.DisplayName}: {ex.Message}";
        }
    }

    public async Task StopAsync()
    {
        _syncTimer.Stop();
        if (_pushCts is not null)
            await _pushCts.CancelAsync();
    }

    // ───────────────────────────── Folders and list ─────────────────────────────

    /// <summary>
    /// Refreshes the folder pane from the cache. Badges update in place; the tree is rebuilt only when
    /// folders were added or removed, and even then the selection never passes through "nothing
    /// selected", which would clear and rebuild the message list (losing scroll position and selection).
    /// </summary>
    private async Task ReloadFoldersAsync()
    {
        if (_account is null) return;
        var folders = await _store.GetFoldersAsync(_account.Id);
        var ordered = OrderAsTree(folders).ToList();

        var sameTree = ordered.Count == Folders.Count && ordered.Select(f => f.Folder.Id).SequenceEqual(Folders.Select(f => f.Id));
        if (sameTree)
        {
            foreach (var (folder, _) in ordered)
                Folders.First(vm => vm.Id == folder.Id).UnreadCount = folder.UnreadCount;
        }
        else
        {
            var selectedId = SelectedFolder?.Id;
            _rebuildingFolders = true;
            try
            {
                Folders.Clear();
                foreach (var (folder, depth) in ordered)
                    Folders.Add(new FolderViewModel(folder, depth));
                SelectedFolder = Folders.FirstOrDefault(f => f.Id == selectedId)
                                 ?? Folders.FirstOrDefault(f => f.Model.Special == SpecialFolder.Inbox)
                                 ?? Folders.FirstOrDefault();
            }
            finally
            {
                _rebuildingFolders = false;
            }
            OnPropertyChanged(nameof(IsDraftsFolder));
        }
        await ReloadMessagesAsync();
    }

    /// <summary>Depth-first order so children sit under their parent, special folders first at the top level.</summary>
    private static IEnumerable<(MailFolder Folder, int Depth)> OrderAsTree(IReadOnlyList<MailFolder> folders)
    {
        var names = folders.Select(f => f.FullName).ToHashSet();
        var byParent = folders.ToLookup(f => f.ParentFullName is not null && names.Contains(f.ParentFullName) ? f.ParentFullName : "");

        IEnumerable<(MailFolder, int)> Walk(string parent, int depth)
        {
            foreach (var f in byParent[parent])   // already sorted by MailStore
            {
                yield return (f, depth);
                foreach (var child in Walk(f.FullName, depth + 1))
                    yield return child;
            }
        }
        return Walk("", 0);
    }

    partial void OnSelectedFolderChanged(FolderViewModel? value)
    {
        if (_rebuildingFolders) return;   // ReloadFoldersAsync reloads once it has restored the selection
        OnPropertyChanged(nameof(IsDraftsFolder));
        // Picking a folder leaves search, as in Outlook.
        if (_activeSearch is not null)
        {
            _activeSearch = null;
            SearchText = "";
        }
        _ = ReloadMessagesAsync();
    }

    partial void OnShowAsConversationsChanged(bool value)
    {
        _settings.ShowAsConversations = value;
        _settings.Save();
        _ = ReloadMessagesAsync();
    }

    partial void OnSearchAllFoldersChanged(bool value)
    {
        _settings.SearchAllFolders = value;
        _settings.Save();
        if (_activeSearch is not null)
            _ = ReloadMessagesAsync();
    }

    /// <summary>One list row: a message, or a conversation headed by its newest message in this folder.</summary>
    private sealed record Row(string Key, IReadOnlyList<MailMessage> InFolder, IReadOnlyList<MailMessage> All, bool IsConversation, string? FolderName);

    private async Task ReloadMessagesAsync()
    {
        if (_account is null) return;
        var folder = SelectedFolder;
        var search = _activeSearch;
        var conversations = ShowAsConversations && search is null;
        var key = search is not null
            ? $"search:{SearchText}:{(SearchAllFolders ? "all" : folder?.Id)}"
            : $"folder:{folder?.Id}:{conversations}";
        FolderTitle = search is not null ? "Search results" : folder?.Name ?? "";

        if (folder is null && search is null)
        {
            Messages.Clear();
            _loadedListKey = null;
            return;
        }

        var generation = ++_reloadGeneration;
        var rows = await LoadRowsAsync(_account.Id, folder, search, conversations);
        // A newer reload started meanwhile (another sync, a new search, the user switched folder):
        // let it win, so two loads never interleave their changes to the list.
        if (generation != _reloadGeneration) return;

        if (_loadedListKey == key)
        {
            MergeRows(rows);
        }
        else
        {
            Messages.Clear();
            foreach (var row in rows)
                Messages.Add(CreateItem(row));
            _loadedListKey = key;
        }
        if (search is not null)
            FolderTitle = rows.Count == 0 ? "No results" : $"Search results ({rows.Count})";
        RegroupHeaders();
    }

    private async Task<List<Row>> LoadRowsAsync(int accountId, FolderViewModel? folder, MailSearchQuery? search, bool conversations)
    {
        if (search is not null)
        {
            var scope = SearchAllFolders ? null : folder?.Id;
            var names = Folders.ToDictionary(f => f.Id, f => f.Name);
            return (await _store.SearchAsync(accountId, scope, search))
                .Select(m => new Row("m:" + m.Id, [m], [m], false, names.GetValueOrDefault(m.FolderId)))
                .ToList();
        }

        var messages = await _store.GetMessagesAsync(folder!.Id);
        if (!conversations)
            return messages.Select(m => new Row("m:" + m.Id, [m], [m], false, null)).ToList();

        // Group by conversation; each thread also pulls in its messages from other folders (your
        // replies in Sent), which Outlook shows in the conversation too.
        var byThread = messages.GroupBy(m => m.ThreadId).ToList();
        var everywhere = (await _store.GetThreadMessagesAsync(accountId, byThread.Select(g => g.Key).ToList()))
            .ToLookup(m => m.ThreadId);
        return byThread
            .Select(g =>
            {
                var inFolder = g.OrderByDescending(m => m.Date).ToList();
                var all = everywhere[g.Key].OrderByDescending(m => m.Date).ToList();
                return new Row("t:" + g.Key, inFolder, all.Count > 0 ? all : inFolder, true, null);
            })
            .OrderByDescending(r => r.InFolder[0].Date)
            .ToList();
    }

    private MessageItemViewModel CreateItem(Row row)
    {
        var item = new MessageItemViewModel(row.InFolder[0]) { IsConversationRow = row.IsConversation };
        ApplyRow(item, row);
        return item;
    }

    private void ApplyRow(MessageItemViewModel item, Row row)
    {
        if (row.IsConversation)
            item.SetThread(row.InFolder, row.All);
        else
            item.Update(row.InFolder[0]);
        item.FolderName = row.FolderName;
        item.Categories = ChipsFor(row.InFolder.SelectMany(m => m.CategoryKeywordList).Distinct());
    }

    /// <summary>
    /// Brings the list in line with the cache without replacing items, so selection, scroll position
    /// and the open message survive a background refresh.
    /// </summary>
    private void MergeRows(IReadOnlyList<Row> rows)
    {
        var freshKeys = rows.Select(r => r.Key).ToHashSet();
        for (var i = Messages.Count - 1; i >= 0; i--)
            if (!freshKeys.Contains(Messages[i].Key))
                Messages.RemoveAt(i);

        var existing = Messages.ToDictionary(m => m.Key);
        for (var i = 0; i < rows.Count; i++)
        {
            if (existing.TryGetValue(rows[i].Key, out var item))
            {
                ApplyRow(item, rows[i]);
                var current = Messages.IndexOf(item);
                if (current != i)
                    Messages.Move(current, i);
            }
            else
            {
                Messages.Insert(i, CreateItem(rows[i]));
            }
        }
    }

    private void RegroupHeaders()
    {
        string? lastGroup = null;
        foreach (var item in Messages)
        {
            item.GroupLabel = MessageItemViewModel.GroupFor(item.LocalDate);
            item.ShowGroupHeader = item.GroupLabel != lastGroup;
            lastGroup = item.GroupLabel;
        }
    }

    /// <summary>Called by the view when the list's multi-selection changes.</summary>
    public void SetSelection(IEnumerable<MessageItemViewModel> items) => _selection = items.ToList();

    /// <summary>The rows an action applies to: the multi-selection, or the focused row.</summary>
    private List<MessageItemViewModel> ActionTargets()
    {
        var targets = _selection.Where(Messages.Contains).ToList();
        if (targets.Count == 0 && SelectedMessage is not null)
            targets.Add(SelectedMessage);
        return targets;
    }

    /// <summary>The messages behind those rows: a conversation row stands for all its messages in this folder.</summary>
    private static List<MailMessage> ModelsOf(IEnumerable<MessageItemViewModel> rows) =>
        rows.SelectMany(r => r.ThreadInFolder).ToList();

    // ───────────────────────────── Search ─────────────────────────────

    [RelayCommand]
    private Task SearchAsync()
    {
        var query = MailSearchQuery.Parse(SearchText);
        if (query.IsEmpty)
            return ClearSearchAsync();
        _activeSearch = query;
        OnPropertyChanged(nameof(IsSearchActive));
        return ReloadMessagesAsync();
    }

    [RelayCommand]
    private Task ClearSearchAsync()
    {
        _activeSearch = null;
        SearchText = "";
        OnPropertyChanged(nameof(IsSearchActive));
        return ReloadMessagesAsync();
    }

    // ───────────────────────────── Reading pane ─────────────────────────────

    partial void OnSelectedMessageChanged(MessageItemViewModel? value)
    {
        OnPropertyChanged(nameof(IsSingleMessageOpen));
        _ = OpenMessageAsync(value);
    }

    private async Task OpenMessageAsync(MessageItemViewModel? item)
    {
        ReadingBody.Clear();
        ConversationCards.Clear();
        IsConversationOpen = false;
        if (item is null) return;

        if (item.IsConversationRow && item.ThreadAll.Count > 1)
        {
            // Outlook's conversation pane: newest first, the newest message and any unread ones open.
            IsConversationOpen = true;
            var folderNames = Folders.ToDictionary(f => f.Id, f => f.Name);
            foreach (var m in item.ThreadAll)
                ConversationCards.Add(new ConversationCardViewModel(m,
                    m.FolderId == SelectedFolder?.Id ? null : folderNames.GetValueOrDefault(m.FolderId), _mail, _store));
            var toOpen = ConversationCards.Where((c, i) => i == 0 || !c.Message.IsRead).ToList();
            foreach (var card in toOpen)
            {
                await card.ExpandAsync();
                if (SelectedMessage != item) return;
            }
        }
        else
        {
            if (!await ReadingBody.LoadAsync(item.Model) || SelectedMessage != item) return;
        }

        // Outlook's default: opening a message (or conversation) in the reading pane marks it read.
        if (!item.IsRead && !item.Model.IsDraft)
            await SetReadAsync([item], true);
    }

    public Task<byte[]> GetAttachmentDataAsync(AttachmentViewModel attachment) =>
        _store.GetAttachmentDataAsync(attachment.Attachment!.Id);

    // ───────────────────────────── Read / flag ─────────────────────────────

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task MarkReadAsync() => SetReadAsync(ActionTargets(), true);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task MarkUnreadAsync() => SetReadAsync(ActionTargets(), false);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task ToggleReadAsync()
    {
        var targets = ActionTargets();
        return SetReadAsync(targets, !targets.All(m => m.IsRead));
    }

    /// <summary>Flags the row's newest message; for a conversation that is what Outlook flags too.</summary>
    [RelayCommand(CanExecute = nameof(HasSelection))]
    private async Task ToggleFlagAsync()
    {
        var targets = ActionTargets();
        var flag = !targets.All(m => m.IsFlagged);
        foreach (var m in targets.Where(m => m.IsFlagged != flag))
        {
            m.IsFlagged = flag;   // optimistic
            try
            {
                await _mail.SetFlaggedAsync(m.Model, flag);
            }
            catch (Exception ex)
            {
                m.IsFlagged = !flag;
                StatusText = $"Couldn't update the flag: {ex.Message}";
            }
        }
    }

    private async Task SetReadAsync(IEnumerable<MessageItemViewModel> rows, bool isRead)
    {
        foreach (var row in rows.ToList())
        {
            var changing = row.ThreadInFolder.Where(m => m.IsRead != isRead).ToList();
            if (changing.Count == 0) continue;
            row.IsRead = isRead;   // optimistic
            AdjustUnread(isRead ? -changing.Count : changing.Count);
            foreach (var m in changing)
            {
                try
                {
                    await _mail.SetReadAsync(m, isRead);
                }
                catch (Exception ex)
                {
                    row.IsRead = !isRead;
                    AdjustUnread(isRead ? 1 : -1);
                    StatusText = $"Couldn't update the message: {ex.Message}";
                }
            }
        }
    }

    private void AdjustUnread(int delta)
    {
        if (SelectedFolder is { } f && _activeSearch is null)
            f.UnreadCount = Math.Max(0, f.UnreadCount + delta);
    }

    // ───────────────────────────── Categories ─────────────────────────────

    private async Task LoadCategoriesAsync()
    {
        if (_account is null) return;
        Categories = await _store.GetCategoriesAsync(_account.Id);
        _categoriesByKeyword = Categories.ToDictionary(c => c.Keyword);
        OnPropertyChanged(nameof(Categories));
    }

    private IReadOnlyList<CategoryChip> ChipsFor(IEnumerable<string> keywords) =>
        keywords.Select(k => _categoriesByKeyword.TryGetValue(k, out var c)
                ? new CategoryChip(c.Name, BrushFor(c.Color), k)
                : new CategoryChip(Category.NameFromKeyword(k), Brushes.Gray, k))   // set by another client
            .ToList();

    public static IBrush BrushFor(string color) =>
        Color.TryParse(color, out var parsed) ? new SolidColorBrush(parsed) : Brushes.Gray;

    /// <summary>True when every selected message already has the category (shown as a check in menus).</summary>
    public bool SelectionHasCategory(Category category)
    {
        var models = ModelsOf(ActionTargets());
        return models.Count > 0 && models.All(m => m.CategoryKeywordList.Contains(category.Keyword));
    }

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private async Task ToggleCategoryAsync(Category category)
    {
        var add = !SelectionHasCategory(category);
        await ApplyCategoryAsync(ActionTargets(), category.Keyword, add);
    }

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private async Task ClearCategoriesAsync()
    {
        var targets = ActionTargets();
        foreach (var keyword in ModelsOf(targets).SelectMany(m => m.CategoryKeywordList).Distinct().ToList())
            await ApplyCategoryAsync(targets, keyword, add: false);
    }

    /// <summary>Creates a category (next colour in the palette) and puts the selection in it.</summary>
    public async Task CreateCategoryAsync(string name)
    {
        if (_account is null || string.IsNullOrWhiteSpace(name)) return;
        var color = NewCategoryColors[Math.Max(0, Categories.Count - Category.Defaults.Length) % NewCategoryColors.Length];
        var category = await _store.AddCategoryAsync(_account.Id, name, color);
        await LoadCategoriesAsync();
        if (HasSelection)
            await ApplyCategoryAsync(ActionTargets(), category.Keyword, add: true);
    }

    private async Task ApplyCategoryAsync(List<MessageItemViewModel> rows, string keyword, bool add)
    {
        var models = ModelsOf(rows);
        if (models.Count == 0) return;
        try
        {
            await _mail.SetCategoryAsync(models, keyword, add);
        }
        catch (Exception ex)
        {
            StatusText = $"Couldn't change the category: {ex.Message}";
            return;
        }
        // The service updated the cache; reflect it on the rows without a full reload.
        foreach (var m in models)
        {
            var set = m.CategoryKeywordList.ToHashSet();
            if (add) set.Add(keyword); else set.Remove(keyword);
            m.CategoryKeywords = string.Join(' ', set.Order(StringComparer.Ordinal));
        }
        foreach (var row in rows)
            row.Categories = ChipsFor(row.ThreadInFolder.SelectMany(m => m.CategoryKeywordList).Distinct());
    }

    // ───────────────────────────── Move / delete / archive ─────────────────────────────

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private async Task DeleteAsync()
    {
        var targets = ActionTargets();
        // Only deleting from Trash is permanent; elsewhere the service moves to Trash (creating it if needed).
        if (SelectedFolder?.Model.Special == SpecialFolder.Trash && _activeSearch is null)
        {
            var count = ModelsOf(targets).Count;
            var question = count == 1
                ? "This message will be permanently deleted. Continue?"
                : $"These {count} messages will be permanently deleted. Continue?";
            if (Confirm is not null && !await Confirm(question))
                return;
        }
        await RemoveFromListAsync(targets, models => _mail.DeleteAsync(models), "delete");
    }

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task ArchiveAsync() =>
        RemoveFromListAsync(ActionTargets(), models => _mail.ArchiveAsync(models), "archive");

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task MoveToAsync(FolderViewModel target) => MoveItemsAsync(ActionTargets(), target);

    /// <summary>Used by drag and drop, which carries its own item list.</summary>
    public Task MoveItemsAsync(IReadOnlyList<MessageItemViewModel> items, FolderViewModel target)
    {
        if (target.Id == SelectedFolder?.Id && _activeSearch is null) return Task.CompletedTask;
        return RemoveFromListAsync(items, models => _mail.MoveAsync(models, target.Model), $"move to {target.Name}");
    }

    /// <summary>Optimistic removal: the rows leave the list at once and come back if the server refuses.</summary>
    private async Task RemoveFromListAsync(IReadOnlyList<MessageItemViewModel> rows, Func<IReadOnlyList<MailMessage>, Task> action, string verb)
    {
        if (rows.Count == 0) return;
        var models = ModelsOf(rows);
        var firstIndex = rows.Select(Messages.IndexOf).Where(i => i >= 0).DefaultIfEmpty(0).Min();
        foreach (var row in rows)
            Messages.Remove(row);
        AdjustUnread(-models.Count(m => !m.IsRead));
        RegroupHeaders();
        // Like Outlook, move the selection to the next message down.
        SelectedMessage = Messages.Count == 0 ? null : Messages[Math.Min(firstIndex, Messages.Count - 1)];

        try
        {
            await action(models);
        }
        catch (Exception ex)
        {
            StatusText = $"Couldn't {verb}: {ex.Message}";
        }
        await ReloadFoldersAsync();
    }

    // ───────────────────────────── Compose and undo send ─────────────────────────────

    private bool CanCompose() => _account is not null;

    [RelayCommand(CanExecute = nameof(CanCompose))]
    private Task NewMailAsync() => StartComposeAsync(ComposeMode.New);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task ReplyAsync() => StartComposeAsync(ComposeMode.Reply);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task ReplyAllAsync() => StartComposeAsync(ComposeMode.ReplyAll);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task ForwardAsync() => StartComposeAsync(ComposeMode.Forward);

    [RelayCommand(CanExecute = nameof(HasSelection))]
    private Task OpenDraftAsync() => IsDraftsFolder ? StartComposeAsync(ComposeMode.EditDraft) : Task.CompletedTask;

    /// <summary>mailto: links in the reading pane.</summary>
    public void ComposeTo(string address)
    {
        if (_account is null) return;
        var seed = new ComposeSeed(address, "", "", "", null, null);
        OpenCompose(new ComposeViewModel(_mail, _account, ComposeMode.New, null, seed, [], null));
    }

    private async Task StartComposeAsync(ComposeMode mode)
    {
        if (_account is null) return;
        MailMessage? original = null;
        var body = "";
        var attachments = new List<OutgoingAttachment>();

        if (mode != ComposeMode.New)
        {
            if (SelectedMessage is not { } item) return;
            // For a conversation, respond to its newest message in this folder.
            var target = item.Model;
            try
            {
                original = ReadingBody.Full?.Id == target.Id ? ReadingBody.Full
                    : ConversationCards.FirstOrDefault(c => c.Message.Id == target.Id)?.Body.Full
                      ?? (target.IsBodyDownloaded ? await _store.GetMessageAsync(target.Id) : await _mail.DownloadBodyAsync(target));
                if (original is null) return;

                body = original.BodyText is { Length: > 0 } text ? text : HtmlText.ToPlainText(original.BodyHtml ?? "");
                // Forwards and drafts carry their files along; replies do not (Outlook's default).
                if (mode is ComposeMode.Forward or ComposeMode.EditDraft)
                    foreach (var a in (await _store.GetAttachmentsAsync(original.Id)).Where(a => !a.IsInline))
                        attachments.Add(new OutgoingAttachment(a.FileName, a.ContentType, await _store.GetAttachmentDataAsync(a.Id)));
            }
            catch (Exception ex)
            {
                StatusText = $"Couldn't open the message: {ex.Message}";
                return;
            }
        }

        var seed = ReplyBuilder.Create(mode, original, body, _account.Email);
        var draftId = mode == ComposeMode.EditDraft ? original?.MessageIdHeader : null;
        OpenCompose(new ComposeViewModel(_mail, _account, mode, original, seed, attachments, draftId));
    }

    private void OpenCompose(ComposeViewModel compose)
    {
        compose.Completed += problem => Dispatcher.UIThread.Post(() =>
        {
            if (problem is not null)
                StatusText = problem;
            _ = ReloadFoldersAsync();   // Sent, Drafts and the replied/forwarded icon changed
        });
        if (_settings.UndoSendSeconds > 0)
            compose.QueueSend = QueueSend;
        ComposeRequested?.Invoke(compose);
    }

    private void QueueSend(ComposeViewModel compose)
    {
        var pending = new PendingSendViewModel(compose, _settings.UndoSendSeconds);
        PendingSends.Add(pending);
        _ = CountDownAndSendAsync(pending);
    }

    private async Task CountDownAndSendAsync(PendingSendViewModel pending)
    {
        try
        {
            while (pending.SecondsLeft > 0)
            {
                await Task.Delay(TimeSpan.FromSeconds(1), pending.Countdown.Token);
                pending.SecondsLeft--;
            }
        }
        catch (OperationCanceledException)
        {
            return;   // undone, or sent early because the app is closing
        }
        PendingSends.Remove(pending);
        await SendPendingAsync(pending);
    }

    private async Task SendPendingAsync(PendingSendViewModel pending)
    {
        StatusText = $"Sending “{pending.Subject}”…";
        if (await pending.Compose.SendNowAsync())
        {
            StatusText = $"Sent “{pending.Subject}”.";
        }
        else
        {
            // Never lose a message: bring it back with the error so the user can retry.
            StatusText = $"Couldn't send “{pending.Subject}”. It's open again so you can retry.";
            ComposeRequested?.Invoke(pending.Compose);
        }
    }

    [RelayCommand]
    private void UndoSend(PendingSendViewModel pending)
    {
        pending.Countdown.Cancel();
        PendingSends.Remove(pending);
        pending.Compose.Reopen();
        pending.Compose.Status = "Sending cancelled. Edit and send again when ready.";
        ComposeRequested?.Invoke(pending.Compose);
    }

    public bool HasPendingSends => PendingSends.Count > 0;

    /// <summary>On exit, send whatever is still waiting rather than dropping it.</summary>
    public async Task FlushPendingSendsAsync()
    {
        foreach (var pending in PendingSends.ToList())
        {
            pending.Countdown.Cancel();
            PendingSends.Remove(pending);
            await pending.Compose.SendNowAsync();
        }
    }
}
