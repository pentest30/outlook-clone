using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class ComposeWindow : Window
{
    /// <summary>Same cap as most providers (Gmail and Outlook.com: 25 MB per message).</summary>
    private const long MaxAttachmentBytes = 25 * 1024 * 1024;

    private bool _closingHandled;

    public ComposeWindow()
    {
        InitializeComponent();
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DropEvent, OnDrop);
    }

    private ComposeViewModel Vm => (ComposeViewModel)DataContext!;

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        // The view model can outlive this window (undo send reopens it in a new one), so unhook on close.
        Vm.CloseRequested += OnCloseRequested;
        if (Vm.FocusBody)
        {
            BodyBox.Focus();
            BodyBox.CaretIndex = 0;   // type the reply above the quoted original
        }
        else
        {
            ToBox.Focus();
        }
    }

    private void OnCloseRequested()
    {
        _closingHandled = true;
        Close();
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        Vm.CloseRequested -= OnCloseRequested;
    }

    protected override async void OnClosing(WindowClosingEventArgs e)
    {
        base.OnClosing(e);
        if (_closingHandled || !Vm.NeedsSaveOnClose)
            return;

        // Save the draft first, then close for real (like Outlook's auto-saved drafts).
        e.Cancel = true;
        _closingHandled = true;
        await Vm.SaveOnCloseAsync();
        Close();
    }

    private async void OnAttachClick(object? sender, RoutedEventArgs e)
    {
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Insert file",
            AllowMultiple = true,
        });
        await AddFilesAsync(files);
    }

    private void OnDragOver(object? sender, DragEventArgs e) =>
        e.DragEffects = e.DataTransfer.Contains(DataFormat.File) ? DragDropEffects.Copy : DragDropEffects.None;

    private async void OnDrop(object? sender, DragEventArgs e)
    {
        if (e.DataTransfer.TryGetFiles() is { } items)
            await AddFilesAsync(items.OfType<IStorageFile>());
    }

    private async Task AddFilesAsync(IEnumerable<IStorageFile> files)
    {
        foreach (var file in files)
        {
            var props = await file.GetBasicPropertiesAsync();
            if (props.Size > MaxAttachmentBytes)
            {
                Vm.Status = $"{file.Name} is larger than 25 MB, which most mail servers reject.";
                continue;
            }
            await using var stream = await file.OpenReadAsync();
            using var buffer = new MemoryStream();
            await stream.CopyToAsync(buffer);
            Vm.AddAttachment(file.Name, buffer.ToArray());
        }
    }
}
