using Avalonia.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class SettingsWindow : Window
{
    public SettingsWindow() => InitializeComponent();

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        if (DataContext is SettingsViewModel vm)
            vm.CloseRequested += Close;
    }
}
