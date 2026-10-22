using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Platform.Storage;
using Avalonia.Media;
using Outlook.App.Controls;
using Outlook.App.ViewModels;
using Outlook.Core.Models;

namespace Outlook.App.Views;

public partial class MainWindow : Window
{
    private static readonly DataFormat<string> DragFormat = DataFormat.CreateStringApplicationFormat("outlook-clone-messages");

    /// <summary>Opening these from mail is how malware spreads; Outlook blocks them too. Saving is still allowed.</summary>
    private static readonly HashSet<string> BlockedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".exe", ".com", ".bat", ".cmd", ".msi", ".msp", ".scr", ".pif", ".cpl", ".js", ".jse", ".vbs", ".vbe",
        ".wsf", ".wsh", ".ps1", ".psm1", ".hta", ".jar", ".lnk", ".reg", ".inf", ".application", ".appref-ms",
    };

    private Point? _dragStart;
    private List<MessageItemViewModel> _dragItems = [];

    public MainWindow()
    {
        InitializeComponent();

        MessageList.SelectionChanged += (_, _) => Vm?.SetSelection(MessageList.SelectedItems?.OfType<MessageItemViewModel>() ?? []);
        MessageList.DoubleTapped += (_, _) => Vm?.OpenDraftCommand.Execute(null);
        MessageList.ContextRequested += OnMessageContextRequested;
        MessageList.AddHandler(PointerPressedEvent, OnListPointerPressed, RoutingStrategies.Tunnel);
        MessageList.AddHandler(PointerMovedEvent, OnListPointerMoved, RoutingStrategies.Tunnel);

        FolderList.AddHandler(DragDrop.DragOverEvent, OnFolderDragOver);
        FolderList.AddHandler(DragDrop.DropEvent, OnFolderDrop);

        AddHandler(MailContent.MailLinkClickedEvent, OnMailLinkClicked);
        AddHandler(KeyDownEvent, OnWindowKeyDown, RoutingStrategies.Tunnel);
    }

    /// <summary>Ctrl+E jumps to search, as in Outlook.</summary>
    private void OnWindowKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.E && e.KeyModifiers == KeyModifiers.Control)
        {
            SearchBox.Focus();
            SearchBox.SelectAll();
            e.Handled = true;
        }
    }

    private MainWindowViewModel? Vm => DataContext as MainWindowViewModel;

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (Vm is not { } vm) return;
        vm.ComposeRequested += compose => new ComposeWindow { DataContext = compose }.Show(this);
        vm.Confirm = ConfirmAsync;
        vm.Reminders.ShowRequested += ShowReminders;
        vm.IsInForeground = () => IsActive && WindowState != WindowState.Minimized;
        _toasts = new ToastManager(this, message =>
        {
            BringToFront();
            _ = vm.ShowMessageAsync(message);
        });
        vm.NewMailArrived += messages => _toasts.ShowNewMail(messages);
        vm.SettingsRequested += settings => new SettingsWindow { DataContext = settings }.ShowDialog(this);
        // Over whichever window asked (Settings is modal, so the form must be modal on top of it).
        vm.AccountSetupRequested += setup => new AccountSetupWindow { DataContext = setup }.ShowDialog(TopWindow());
    }

    private bool _closeAfterFlush;

    /// <summary>Messages still in their undo window are sent before the app exits, never dropped.</summary>
    protected override async void OnClosing(WindowClosingEventArgs e)
    {
        base.OnClosing(e);
        if (_closeAfterFlush || Vm is not { HasPendingSends: true } vm)
            return;
        e.Cancel = true;
        _closeAfterFlush = true;
        vm.StatusText = "Sending queued messages before closing…";
        await vm.FlushPendingSendsAsync();
        Close();
    }

    protected override async void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        if (Vm is { } vm)
            await vm.StopAsync();
    }

    /// <summary>The active window of the app (a dialog if one is open), to own the next dialog.</summary>
    private Window TopWindow() =>
        (Avalonia.Application.Current?.ApplicationLifetime as Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime)
        ?.Windows.LastOrDefault(w => w.IsActive) ?? this;

    /// <summary>The address in the title bar: switch account, add one, or open Settings › Accounts.</summary>
    private void OnAccountButtonClick(object? sender, RoutedEventArgs e)
    {
        if (Vm is not { } vm) return;
        var items = new List<Control>();
        foreach (var row in vm.Accounts)
        {
            items.Add(new MenuItem
            {
                Header = row.Email,
                Icon = row.IsActive ? new TextBlock { Text = "✓" } : null,
                Command = vm.SwitchToCommand,
                CommandParameter = row,
            });
        }
        items.Add(new Separator());
        items.Add(new MenuItem { Header = "Add account…", Command = vm.AddAccountCommand });
        items.Add(new MenuItem { Header = "Manage accounts…", Command = vm.ManageAccountsCommand });
        new ContextMenu { ItemsSource = items, Placement = PlacementMode.BottomEdgeAlignedRight }.Open(AccountButton);
    }

    private ToastManager? _toasts;
    private WindowState _lastShownState = WindowState.Maximized;

    /// <summary>Restore from minimized (to how it was) and focus, for alerts and the tray icon.</summary>
    public void BringToFront()
    {
        if (!IsVisible) Show();
        if (WindowState == WindowState.Minimized) WindowState = _lastShownState;
        Activate();
    }

    protected override void OnPropertyChanged(Avalonia.AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == WindowStateProperty && WindowState != WindowState.Minimized)
            _lastShownState = WindowState;
    }

    private RemindersWindow? _reminders;

    private void ShowReminders()
    {
        if (Vm is not { } vm) return;
        if (_reminders is null)
        {
            _reminders = new RemindersWindow { DataContext = vm.Reminders };
            _reminders.Closed += (_, _) => _reminders = null;
            // No owner: owned windows are hidden while the owner is minimized, when reminders matter most.
            _reminders.Show();
        }
        else
        {
            _reminders.Activate();
        }
    }

    // ───────────────────────────── Context menus ─────────────────────────────

    private void OnMessageContextRequested(object? sender, ContextRequestedEventArgs e)
    {
        if (Vm is not { HasSelection: true } vm) return;
        var menu = new ContextMenu
        {
            ItemsSource = new Control[]
            {
                Item("Reply", vm.ReplyCommand, "Ctrl+R"),
                Item("Reply all", vm.ReplyAllCommand, "Ctrl+Shift+R"),
                Item("Forward", vm.ForwardCommand, "Ctrl+F"),
                new Separator(),
                Item("Mark as read", vm.MarkReadCommand, "Ctrl+Q"),
                Item("Mark as unread", vm.MarkUnreadCommand, "Ctrl+U"),
                Item("Flag / unflag", vm.ToggleFlagCommand, "Insert"),
                new Separator(),
                MoveToMenu(vm),
                new MenuItem { Header = "Categorize", ItemsSource = CategoryItems(vm) },
                Item("Archive", vm.ArchiveCommand, "Back"),
                Item("Delete", vm.DeleteCommand, "Delete"),
            },
        };
        menu.Open(MessageList);
        e.Handled = true;
    }

    private void OnMoveButtonClick(object? sender, RoutedEventArgs e)
    {
        if (Vm is not { } vm) return;
        var menu = new ContextMenu { ItemsSource = MoveTargets(vm), Placement = PlacementMode.BottomEdgeAlignedLeft };
        menu.Open(MoveButton);
    }

    private void OnCategorizeButtonClick(object? sender, RoutedEventArgs e)
    {
        if (Vm is not { } vm) return;
        new ContextMenu { ItemsSource = CategoryItems(vm), Placement = PlacementMode.BottomEdgeAlignedLeft }.Open(CategorizeButton);
    }

    /// <summary>Outlook's Categorize menu: a colour square per category, checked when the selection has it.</summary>
    private List<Control> CategoryItems(MainWindowViewModel vm)
    {
        var items = new List<Control>();
        foreach (var category in vm.Categories)
        {
            var isSet = vm.SelectionHasCategory(category);
            items.Add(new MenuItem
            {
                Header = category.Name,
                Icon = new Border
                {
                    Width = 14, Height = 14, CornerRadius = new CornerRadius(3),
                    Background = MainWindowViewModel.BrushFor(category.Color),
                    Child = new TextBlock
                    {
                        Text = isSet ? "✓" : "", Foreground = Brushes.White, FontSize = 10,
                        HorizontalAlignment = HorizontalAlignment.Center, VerticalAlignment = VerticalAlignment.Center,
                    },
                },
                Command = vm.ToggleCategoryCommand,
                CommandParameter = category,
            });
        }
        items.Add(new Separator());
        var create = new MenuItem { Header = "New category…" };
        create.Click += async (_, _) =>
        {
            if (await PromptAsync("New category", "Name") is { Length: > 0 } name)
                await vm.CreateCategoryAsync(name);
        };
        items.Add(create);
        items.Add(new MenuItem { Header = "Clear all categories", Command = vm.ClearCategoriesCommand });
        return items;
    }

    private static MenuItem MoveToMenu(MainWindowViewModel vm) => new() { Header = "Move to", ItemsSource = MoveTargets(vm) };

    private static List<MenuItem> MoveTargets(MainWindowViewModel vm) =>
        vm.Folders
            .Where(f => f.Id != vm.SelectedFolder?.Id)
            .Select(f => new MenuItem
            {
                Header = $"{f.Icon}  {f.Name}",
                Padding = new Thickness(f.Indent.Left - 4, 4, 12, 4),
                Command = vm.MoveToCommand,
                CommandParameter = f,
            })
            .ToList();

    private static MenuItem Item(string header, System.Windows.Input.ICommand command, string gesture) =>
        new() { Header = header, Command = command, InputGesture = KeyGesture.Parse(gesture) };

    // ───────────────────────────── Drag messages onto folders ─────────────────────────────

    private void OnListPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        _dragStart = e.GetCurrentPoint(MessageList).Properties.IsLeftButtonPressed ? e.GetPosition(MessageList) : null;
    }

    private async void OnListPointerMoved(object? sender, PointerEventArgs e)
    {
        if (_dragStart is not { } start || Vm is not { } vm) return;
        if (!e.GetCurrentPoint(MessageList).Properties.IsLeftButtonPressed)
        {
            _dragStart = null;
            return;
        }
        var delta = e.GetPosition(MessageList) - start;
        if (Math.Abs(delta.X) < 8 && Math.Abs(delta.Y) < 8) return;
        _dragStart = null;

        // The press already selected the item under the pointer, so the selection is what moves.
        _dragItems = MessageList.SelectedItems?.OfType<MessageItemViewModel>().ToList() ?? [];
        if (_dragItems.Count == 0) return;

        var data = new DataTransfer();
        data.Add(DataTransferItem.Create(DragFormat, _dragItems.Count.ToString()));
        await DragDrop.DoDragDropAsync(e, data, DragDropEffects.Move);
    }

    private void OnFolderDragOver(object? sender, DragEventArgs e)
    {
        var target = FolderAt(e);
        e.DragEffects = e.DataTransfer.Contains(DragFormat) && target is not null && target.Id != Vm?.SelectedFolder?.Id
            ? DragDropEffects.Move
            : DragDropEffects.None;
    }

    private async void OnFolderDrop(object? sender, DragEventArgs e)
    {
        if (!e.DataTransfer.Contains(DragFormat) || FolderAt(e) is not { } target || Vm is not { } vm) return;
        var items = _dragItems;
        _dragItems = [];
        await vm.MoveItemsAsync(items, target);
    }

    private static FolderViewModel? FolderAt(DragEventArgs e) => (e.Source as StyledElement)?.DataContext as FolderViewModel;

    // ───────────────────────────── Links and attachments ─────────────────────────────

    private async void OnMailLinkClicked(object? sender, MailLinkEventArgs e)
    {
        var link = e.Link;
        if (link.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
        {
            var address = Uri.UnescapeDataString(link[7..].Split('?')[0]);
            Vm?.ComposeTo(address);
        }
        else if (Uri.TryCreate(link, UriKind.Absolute, out var uri) && uri.Scheme is "http" or "https")
        {
            await Launcher.LaunchUriAsync(uri);
        }
    }

    private void OnAttachmentClick(object? sender, RoutedEventArgs e)
    {
        if (sender is not Button { DataContext: AttachmentViewModel attachment } button) return;
        var canOpen = !BlockedExtensions.Contains(Path.GetExtension(attachment.FileName));
        var open = new MenuItem { Header = canOpen ? "Open" : "Open (blocked: this file type can run code)", IsEnabled = canOpen };
        open.Click += async (_, _) => await OpenAttachmentAsync(attachment);
        var save = new MenuItem { Header = "Save as…" };
        save.Click += async (_, _) => await SaveAttachmentAsync(attachment);
        new ContextMenu { ItemsSource = new[] { open, save }, Placement = PlacementMode.BottomEdgeAlignedLeft }.Open(button);
    }

    private async Task OpenAttachmentAsync(AttachmentViewModel attachment)
    {
        if (Vm is not { } vm) return;
        // A fresh folder per open, so two attachments with the same name never collide.
        var dir = Path.Combine(Path.GetTempPath(), "OutlookClone", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, attachment.FileName);
        await File.WriteAllBytesAsync(path, await vm.GetAttachmentDataAsync(attachment));
        await Launcher.LaunchFileInfoAsync(new FileInfo(path));
    }

    private async Task SaveAttachmentAsync(AttachmentViewModel attachment)
    {
        if (Vm is not { } vm) return;
        var file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save attachment",
            SuggestedFileName = attachment.FileName,
            ShowOverwritePrompt = true,
        });
        if (file is null) return;
        await using var stream = await file.OpenWriteAsync();
        await stream.WriteAsync(await vm.GetAttachmentDataAsync(attachment));
    }

    // ───────────────────────────── Dialogs ─────────────────────────────

    /// <summary>A one-line text prompt; returns null when cancelled.</summary>
    private async Task<string?> PromptAsync(string title, string label)
    {
        string? result = null;
        var dialog = new Window
        {
            Title = title,
            Width = 380,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false,
        };
        var input = new TextBox { Watermark = label };
        var ok = new Button { Content = "OK", Classes = { "primary" }, IsDefault = true, MinWidth = 80 };
        var cancel = new Button { Content = "Cancel", IsCancel = true, MinWidth = 80 };
        ok.Click += (_, _) => { result = input.Text?.Trim(); dialog.Close(); };
        cancel.Click += (_, _) => dialog.Close();
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 16,
            Children =
            {
                input,
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { ok, cancel },
                },
            },
        };
        dialog.Opened += (_, _) => input.Focus();
        await dialog.ShowDialog(this);
        return result;
    }

    private async Task<bool> ConfirmAsync(string question)
    {
        var result = false;
        var dialog = new Window
        {
            Title = "Outlook Clone",
            Width = 420,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            ShowInTaskbar = false,
        };
        var yes = new Button { Content = "Yes", Classes = { "primary" }, IsDefault = true, MinWidth = 80 };
        var no = new Button { Content = "No", IsCancel = true, MinWidth = 80 };
        yes.Click += (_, _) => { result = true; dialog.Close(); };
        no.Click += (_, _) => dialog.Close();
        dialog.Content = new StackPanel
        {
            Margin = new Thickness(20),
            Spacing = 16,
            Children =
            {
                new TextBlock { Text = question, TextWrapping = Avalonia.Media.TextWrapping.Wrap },
                new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { yes, no },
                },
            },
        };
        await dialog.ShowDialog(this);
        return result;
    }
}
