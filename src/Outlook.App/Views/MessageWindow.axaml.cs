using Avalonia.Controls;
using Outlook.App.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class MessageWindow : Window
{
    public MessageWindow()
    {
        InitializeComponent();
        AddHandler(MailContent.MailLinkClickedEvent, async (_, e) =>
        {
            if (DataContext is MessageWindowViewModel vm)
                await MailActions.FollowLinkAsync(this, e.Link, vm.ComposeTo);
        });
    }

    protected override void OnDataContextChanged(EventArgs e)
    {
        base.OnDataContextChanged(e);
        if (DataContext is MessageWindowViewModel vm)
            vm.CloseRequested += Close;
    }
}
