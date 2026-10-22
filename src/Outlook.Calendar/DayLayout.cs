namespace Outlook.Calendar;

/// <summary>
/// Places overlapping appointments side by side, as Outlook's day view does: appointments that
/// overlap (directly or through a chain) form a cluster, each gets the first free lane, and every
/// member of a cluster shares the cluster's lane count so the column widths line up.
/// </summary>
public static class DayLayout
{
    public readonly record struct Slot(int Index, int Lane, int LaneCount);

    public static IReadOnlyList<Slot> Arrange(IReadOnlyList<(DateTimeOffset Start, DateTimeOffset End)> items)
    {
        var order = Enumerable.Range(0, items.Count)
            .OrderBy(i => items[i].Start).ThenByDescending(i => items[i].End - items[i].Start)
            .ToList();
        var slots = new Slot[items.Count];
        var cluster = new List<int>();
        var laneEnds = new List<DateTimeOffset>();   // when each lane of the current cluster frees up
        var clusterEnd = DateTimeOffset.MinValue;

        void CloseCluster()
        {
            foreach (var i in cluster)
                slots[i] = slots[i] with { LaneCount = laneEnds.Count };
            cluster.Clear();
            laneEnds.Clear();
        }

        foreach (var i in order)
        {
            var (start, end) = items[i];
            if (end <= start) end = start.AddMinutes(1);   // zero-length items still need a sliver
            if (cluster.Count > 0 && start >= clusterEnd)
                CloseCluster();

            var lane = laneEnds.FindIndex(free => free <= start);
            if (lane < 0)
            {
                lane = laneEnds.Count;
                laneEnds.Add(end);
            }
            else
            {
                laneEnds[lane] = end;
            }
            slots[i] = new Slot(i, lane, 0);
            cluster.Add(i);
            if (end > clusterEnd || cluster.Count == 1) clusterEnd = cluster.Count == 1 ? end : Max(clusterEnd, end);
        }
        if (cluster.Count > 0)
            CloseCluster();
        return slots;
    }

    private static DateTimeOffset Max(DateTimeOffset a, DateTimeOffset b) => a > b ? a : b;
}
