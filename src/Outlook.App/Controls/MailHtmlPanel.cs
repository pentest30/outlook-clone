using Avalonia;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Outlook.Core.Text;
using TheArtOfDev.HtmlRenderer.Avalonia;
using TheArtOfDev.HtmlRenderer.Core.Entities;

namespace Outlook.App.Controls;

/// <summary>
/// Reading-pane HTML view (scrolls itself). On top of HtmlPanel it serves inline cid: pictures from
/// the cached attachments and refuses anything from the web until AllowRemoteContent is set.
/// </summary>
public class MailHtmlPanel : HtmlPanel
{
    public static readonly StyledProperty<bool> AllowRemoteContentProperty =
        AvaloniaProperty.Register<MailHtmlPanel, bool>(nameof(AllowRemoteContent));

    public static readonly StyledProperty<Func<string, byte[]?>?> InlineImageResolverProperty =
        AvaloniaProperty.Register<MailHtmlPanel, Func<string, byte[]?>?>(nameof(InlineImageResolver));

    public bool AllowRemoteContent
    {
        get => GetValue(AllowRemoteContentProperty);
        set => SetValue(AllowRemoteContentProperty, value);
    }

    /// <summary>Maps a Content-ID (without the "cid:" prefix) to the picture's bytes.</summary>
    public Func<string, byte[]?>? InlineImageResolver
    {
        get => GetValue(InlineImageResolverProperty);
        set => SetValue(InlineImageResolverProperty, value);
    }

    // Use HtmlPanel's control theme.
    protected override Type StyleKeyOverride => typeof(HtmlPanel);

    public MailHtmlPanel()
    {
        ImageLoad += (_, e) => MailContent.LoadImage(e.Event, AllowRemoteContent, InlineImageResolver);
        StylesheetLoad += (_, e) => MailContent.LoadStylesheet(e.Event, AllowRemoteContent);
        LinkClicked += (_, e) => MailContent.ForwardLink(this, e.Event);
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == AllowRemoteContentProperty || change.Property == InlineImageResolverProperty)
            MailContent.Rerender(this);
    }
}

/// <summary>
/// The same mail rules for a body that sizes to its content, for the stacked cards of a conversation
/// (an HtmlPanel would scroll inside the stack's own scroll viewer).
/// </summary>
public class MailHtmlLabel : HtmlLabel
{
    public static readonly StyledProperty<bool> AllowRemoteContentProperty =
        MailHtmlPanel.AllowRemoteContentProperty.AddOwner<MailHtmlLabel>();

    public static readonly StyledProperty<Func<string, byte[]?>?> InlineImageResolverProperty =
        MailHtmlPanel.InlineImageResolverProperty.AddOwner<MailHtmlLabel>();

    public bool AllowRemoteContent
    {
        get => GetValue(AllowRemoteContentProperty);
        set => SetValue(AllowRemoteContentProperty, value);
    }

    public Func<string, byte[]?>? InlineImageResolver
    {
        get => GetValue(InlineImageResolverProperty);
        set => SetValue(InlineImageResolverProperty, value);
    }

    protected override Type StyleKeyOverride => typeof(HtmlLabel);

    public MailHtmlLabel()
    {
        AutoSizeHeightOnly = true;
        ImageLoad += (_, e) => MailContent.LoadImage(e.Event, AllowRemoteContent, InlineImageResolver);
        StylesheetLoad += (_, e) => MailContent.LoadStylesheet(e.Event, AllowRemoteContent);
        LinkClicked += (_, e) => MailContent.ForwardLink(this, e.Event);
    }

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == AllowRemoteContentProperty || change.Property == InlineImageResolverProperty)
            MailContent.Rerender(this);
    }
}

/// <summary>A link clicked in a mail body. Bubbles, so one window-level handler covers every body view.</summary>
public class MailLinkEventArgs(string link) : RoutedEventArgs(MailContent.MailLinkClickedEvent)
{
    public string Link { get; } = link;
}

public static class MailContent
{
    public static readonly RoutedEvent<MailLinkEventArgs> MailLinkClickedEvent =
        RoutedEvent.Register<HtmlControl, MailLinkEventArgs>("MailLinkClicked", RoutingStrategies.Bubble);

    /// <summary>Never let the renderer navigate on its own; hand the link to the window instead.</summary>
    internal static void ForwardLink(HtmlControl source, HtmlLinkClickedEventArgs args)
    {
        args.Handled = true;
        if (!string.IsNullOrWhiteSpace(args.Link))
            source.RaiseEvent(new MailLinkEventArgs(args.Link.Trim()));
    }

    internal static void LoadImage(HtmlImageLoadEventArgs args, bool allowRemote, Func<string, byte[]?>? resolver)
    {
        var src = args.Src ?? "";
        if (src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase))
        {
            args.Handled = true;
            var bytes = resolver?.Invoke(src[4..].Trim('<', '>'));
            if (bytes is null)
            {
                args.Callback();
                return;
            }
            try
            {
                args.Callback(new Bitmap(new MemoryStream(bytes)));
            }
            catch (Exception)
            {
                args.Callback();   // not a picture format we can decode
            }
        }
        else if (EmailHtml.IsRemote(src) && !allowRemote)
        {
            args.Handled = true;
            args.Callback();
        }
    }

    internal static void LoadStylesheet(HtmlStylesheetLoadEventArgs args, bool allowRemote)
    {
        if (EmailHtml.IsRemote(args.Src) && !allowRemote)
            args.SetStyleSheet = "";
    }

    /// <summary>Pictures are resolved during layout, so re-render once the rules change.</summary>
    internal static void Rerender(HtmlControl control)
    {
        var html = control.Text;
        control.Text = "";
        control.Text = html;
    }
}
