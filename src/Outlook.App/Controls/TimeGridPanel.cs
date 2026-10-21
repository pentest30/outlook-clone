using Avalonia;
using Avalonia.Controls;
using Outlook.App.ViewModels;

namespace Outlook.App.Controls;

/// <summary>
/// Lays out one day column of the time grid: each child's DataContext is an EventBlockViewModel,
/// placed vertically by its minutes and horizontally by its lane among overlapping appointments.
/// </summary>
public class TimeGridPanel : Panel
{
    public static readonly StyledProperty<double> HourHeightProperty =
        AvaloniaProperty.Register<TimeGridPanel, double>(nameof(HourHeight), CalendarViewModel.HourHeight);

    public double HourHeight
    {
        get => GetValue(HourHeightProperty);
        set => SetValue(HourHeightProperty, value);
    }

    /// <summary>Shortest drawn height, so a 5-minute item can still show its title.</summary>
    private const double MinBlockHeight = 18;
    private const double Gap = 2;

    protected override Size MeasureOverride(Size availableSize)
    {
        var width = double.IsInfinity(availableSize.Width) ? 200 : availableSize.Width;
        foreach (var child in Children)
        {
            if (child.DataContext is EventBlockViewModel block)
                child.Measure(new Size(Math.Max(0, width / block.LaneCount - Gap), BlockHeight(block)));
        }
        return new Size(width, 24 * HourHeight);
    }

    protected override Size ArrangeOverride(Size finalSize)
    {
        foreach (var child in Children)
        {
            if (child.DataContext is not EventBlockViewModel block) continue;
            var laneWidth = finalSize.Width / Math.Max(1, block.LaneCount);
            var top = block.StartMinute / 60 * HourHeight;
            child.Arrange(new Rect(block.Lane * laneWidth, top, Math.Max(0, laneWidth - Gap), BlockHeight(block)));
        }
        return finalSize;
    }

    private double BlockHeight(EventBlockViewModel block) =>
        Math.Max(MinBlockHeight, (block.EndMinute - block.StartMinute) / 60 * HourHeight - 1);
}
