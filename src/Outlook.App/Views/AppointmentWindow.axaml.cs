using Avalonia.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class AppointmentWindow : Window
{
    public AppointmentWindow() => InitializeComponent();

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        if (DataContext is AppointmentEditorViewModel vm)
            vm.CloseRequested += Close;
        TitleBox.Focus();
    }
}
