using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>A positioned tree node.</summary>
public sealed record TreeNodeBox(string Id, string Label, string? Detail, int Depth, Box Box, string? ParentId);

/// <summary>
/// Top-down tidy tree layout: leaves take consecutive slots left to right, each parent is
/// centered over its children, levels are evenly spaced. Deterministic and pure.
/// </summary>
public static class TreeLayout
{
    /// <summary>Lays out the tree inside <paramref name="area"/>; null when it cannot fit at <paramref name="minNodeWidth"/>.</summary>
    public static IReadOnlyList<TreeNodeBox>? Layout(SpecNode root, Box area, float nodeHeight, float minNodeWidth, float maxNodeWidth, float gap)
    {
        ArgumentNullException.ThrowIfNull(root);
        var leaves = CountLeaves(root);
        var depth = Depth(root);
        var nodeWidth = Math.Min(maxNodeWidth, (area.Width - (gap * (leaves - 1))) / leaves);
        if (nodeWidth < minNodeWidth)
            return null;

        var levelGap = depth <= 1 ? 0 : Math.Max(gap, (area.Height - (depth * nodeHeight)) / (depth - 1));
        levelGap = Math.Min(levelGap, nodeHeight * 1.6f);
        var totalHeight = (depth * nodeHeight) + ((depth - 1) * levelGap);
        if (totalHeight > area.Height + 0.5f)
            return null;

        var usedWidth = (leaves * nodeWidth) + ((leaves - 1) * gap);
        var startX = area.Left + ((area.Width - usedWidth) / 2f);
        var result = new List<TreeNodeBox>();
        var nextLeaf = 0;
        Place(root, "n1", null, 0);
        return result;

        float Place(SpecNode node, string path, string? parentId, int level)
        {
            var id = node.Id ?? path;
            float center;
            var index = result.Count;
            result.Add(null!);
            if (node.Children.Count == 0)
            {
                center = startX + (nextLeaf * (nodeWidth + gap)) + (nodeWidth / 2f);
                nextLeaf++;
            }
            else
            {
                var centers = node.Children.Select((child, i) => Place(child, $"{path}.{i + 1}", id, level + 1)).ToList();
                center = (centers[0] + centers[^1]) / 2f;
            }
            var top = area.Top + (level * (nodeHeight + levelGap));
            result[index] = new TreeNodeBox(id, node.Label, node.Detail, level, LayoutEngine.R(Box.FromSize(center - (nodeWidth / 2f), top, nodeWidth, nodeHeight)), parentId);
            return center;
        }
    }

    /// <summary>Number of leaves.</summary>
    public static int CountLeaves(SpecNode node) => node.Children.Count == 0 ? 1 : node.Children.Sum(CountLeaves);

    /// <summary>Number of levels.</summary>
    public static int Depth(SpecNode node) => 1 + (node.Children.Count == 0 ? 0 : node.Children.Max(Depth));
}
