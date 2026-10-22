using Avalonia.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class AccountSetupWindow : Window
{
    public AccountSetupWindow() => InitializeComponent();

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        if (DataContext is AccountSetupViewModel vm)
            vm.CloseRequested += Close;
    }
}
