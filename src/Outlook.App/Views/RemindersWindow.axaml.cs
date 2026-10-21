using System.Collections.Specialized;
using Avalonia.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class RemindersWindow : Window
{
    public RemindersWindow() => InitializeComponent();

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        if (DataContext is RemindersViewModel vm)
            vm.Items.CollectionChanged += OnItemsChanged;
    }

    protected override void OnClosed(EventArgs e)
    {
        base.OnClosed(e);
        if (DataContext is RemindersViewModel vm)
            vm.Items.CollectionChanged -= OnItemsChanged;
    }

    /// <summary>The window goes away once every reminder is dismissed or snoozed.</summary>
    private void OnItemsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (sender is System.Collections.ICollection { Count: 0 })
            Close();
    }
}
