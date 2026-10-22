using Avalonia.Controls;
using Avalonia.Interactivity;
using Outlook.App.ViewModels;

namespace Outlook.App.Views;

public partial class MessageBodyExtras : UserControl
{
    public MessageBodyExtras() => InitializeComponent();

    private void OnAttachmentClick(object? sender, RoutedEventArgs e)
    {
        if (sender is Button { DataContext: AttachmentViewModel attachment } button)
            MailActions.ShowAttachmentMenu(button, attachment);
    }
}
