using System.Globalization;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.Core.Diagram;

/// <summary>A laid-out diagram: elements to create, text keys to fit, and routing notes.</summary>
public sealed class DiagramPlan
{
    /// <summary>Elements in creation order (backgrounds, nodes, connectors, labels).</summary>
    public required IReadOnlyList<PlannedElement> Elements { get; init; }

    /// <summary>Keys of text-bearing elements that must fit their boxes.</summary>
    public required IReadOnlyList<string> TextKeys { get; init; }

    /// <summary>Node id to element key.</summary>
    public required IReadOnlyDictionary<string, string> NodeKeys { get; init; }

    /// <summary>Layout notes: crossings, long edges, routing limits.</summary>
    public required IReadOnlyList<string> Notes { get; init; }

    /// <summary>Estimated connector crossings between adjacent ranks (flowchart, swimlane, architecture).</summary>
    public int Crossings { get; init; }
}

/// <summary>
/// Deterministic diagram layouts. Flowcharts and swimlanes use a layered layout: back edges are
/// found by depth-first search, ranks are longest paths, and nodes in a rank are ordered by the
/// barycenter of their neighbors (two sweeps). Connector routing itself is PowerPoint's: elbow
/// connectors are glued to connection sites and do not avoid obstacles. Pure.
/// </summary>
public static class DiagramLayout
{
    private const float LabelHeight = 18f;

    /// <summary>Lays out a diagram inside an area.</summary>
    public static DiagramPlan Layout(DiagramSpec spec, Box area, ResolvedProfile profile)
    {
        ArgumentNullException.ThrowIfNull(spec);
        ArgumentNullException.ThrowIfNull(profile);
        var builder = new Builder(spec, profile);
        switch (spec.Type)
        {
            case "flowchart":
                builder.Layered(area, spec.Direction == "right", lanes: null);
                break;
            case "swimlane":
                builder.Swimlane(area);
                break;
            case "matrix":
                builder.Matrix(area);
                break;
            case "hub-spoke":
                builder.HubSpoke(area);
                break;
            case "architecture":
                builder.Architecture(area);
                break;
            default:
                throw new ArgumentException($"Unknown diagram type '{spec.Type}'.");
        }
        return builder.Build();
    }

    /// <summary>Ranks nodes by longest path after reversing back edges; returns ranks and the back edges.</summary>
    public static (IReadOnlyDictionary<string, int> Ranks, IReadOnlySet<(string From, string To)> BackEdges) Rank(
        IReadOnlyList<string> nodes, IReadOnlyList<(string From, string To)> edges)
    {
        ArgumentNullException.ThrowIfNull(nodes);
        ArgumentNullException.ThrowIfNull(edges);
        var outgoing = nodes.ToDictionary(node => node, _ => new List<string>(), StringComparer.Ordinal);
        foreach (var (from, to) in edges)
            outgoing[from].Add(to);

        // Depth-first search in node order; an edge to a node on the stack is a back edge.
        var state = nodes.ToDictionary(node => node, _ => 0, StringComparer.Ordinal);
        var back = new HashSet<(string, string)>();
        void Visit(string node)
        {
            state[node] = 1;
            foreach (var next in outgoing[node])
            {
                if (state[next] == 1)
                    back.Add((node, next));
                else if (state[next] == 0)
                    Visit(next);
            }
            state[node] = 2;
        }
        var incoming = edges.Select(edge => edge.To).ToHashSet(StringComparer.Ordinal);
        foreach (var node in nodes.Where(node => !incoming.Contains(node)).Concat(nodes))
        {
            if (state[node] == 0)
                Visit(node);
        }

        var forward = edges.Where(edge => !back.Contains(edge)).ToList();
        var rank = nodes.ToDictionary(node => node, _ => 0, StringComparer.Ordinal);
        // Longest path by relaxation (graph is acyclic after removing back edges).
        for (int pass = 0; pass < nodes.Count; pass++)
        {
            bool changed = false;
            foreach (var (from, to) in forward)
            {
                if (rank[to] < rank[from] + 1)
                {
                    rank[to] = rank[from] + 1;
                    changed = true;
                }
            }
            if (!changed)
                break;
        }
        return (rank, back);
    }

    /// <summary>Counts crossings between edges joining adjacent ranks, given positions within ranks.</summary>
    public static int CountCrossings(IReadOnlyList<(string From, string To)> edges, IReadOnlyDictionary<string, int> ranks, IReadOnlyDictionary<string, float> positions)
    {
        ArgumentNullException.ThrowIfNull(edges);
        var adjacent = edges.Where(edge => ranks[edge.To] == ranks[edge.From] + 1).ToList();
        int crossings = 0;
        for (int i = 0; i < adjacent.Count; i++)
        {
            for (int j = i + 1; j < adjacent.Count; j++)
            {
                var a = adjacent[i];
                var b = adjacent[j];
                if (ranks[a.From] != ranks[b.From] || a.From == b.From || a.To == b.To)
                    continue;
                if ((positions[a.From] - positions[b.From]) * (positions[a.To] - positions[b.To]) < 0)
                    crossings++;
            }
        }
        return crossings;
    }

    private sealed class Builder(DiagramSpec spec, ResolvedProfile profile)
    {
        private readonly List<PlannedElement> _backgrounds = [];
        private readonly List<PlannedElement> _nodes = [];
        private readonly List<PlannedElement> _connectors = [];
        private readonly List<PlannedElement> _labels = [];
        private readonly List<string> _textKeys = [];
        private readonly Dictionary<string, string> _nodeKeys = new(StringComparer.Ordinal);
        private readonly Dictionary<string, Box> _boxes = new(StringComparer.Ordinal);
        private readonly List<string> _notes = [];
        private int _crossings;

        private ResolvedComponent Node => profile.Component("node");

        private const string Component = "diagram@1";

        public DiagramPlan Build() => new()
        {
            Elements = [.. _backgrounds, .. _nodes, .. _connectors, .. _labels],
            TextKeys = _textKeys,
            NodeKeys = _nodeKeys,
            Notes = _notes,
            Crossings = _crossings,
        };

        private Dictionary<string, string> Tags(string kind, string value) => new(StringComparer.Ordinal)
        {
            ["PPTMCP_DIAGRAM"] = spec.Id,
            [kind] = value,
        };

        private void AddNode(DiagramNode node, Box box, string? fillOverride = null)
        {
            var key = $"node-{node.Id}";
            var style = Node;
            var fill = fillOverride ?? (node.Tone is { } tone ? profile.Color(tone) : style.Fill);
            var textColor = Contrast(fill);
            var paragraphs = new List<PlannedParagraph>
            {
                new() { Text = node.Label, Size = style.HeadingSize, Font = profile.HeadingFont, Color = textColor, Bold = true, Align = "center" },
            };
            if (node.Detail is not null)
                paragraphs.Add(new PlannedParagraph { Text = node.Detail, Size = style.BodySize, Font = profile.BodyFont, Color = textColor, Align = "center" });
            _nodes.Add(new PlannedElement
            {
                Key = key,
                Role = "diagram-node",
                Type = "shape",
                Geometry = node.Shape switch
                {
                    "rounded" => "rounded-rectangle",
                    "decision" => "diamond",
                    "terminator" => "terminator",
                    "data" => "parallelogram",
                    "document" => "document",
                    "database" => "cylinder",
                    "ellipse" => "oval",
                    "hexagon" => "hexagon",
                    _ => "rectangle",
                },
                Radius = node.Shape == "rounded" ? style.Radius : 0,
                Box = LayoutEngine.R(box),
                Fill = fill,
                LineColor = style.Border,
                LineWidth = style.BorderWidth,
                Paragraphs = paragraphs,
                VAlign = "middle",
                Padding = node.Shape == "decision" ? [box.Width * 0.18f, 2, box.Width * 0.18f, 2] : [6, 3, 6, 3],
                Component = Component,
                Tags = new Dictionary<string, string>(Tags(DeckRoles.NodeTag, $"{spec.Id}:{node.Id}"), StringComparer.Ordinal),
            });
            _textKeys.Add(key);
            _nodeKeys[node.Id] = key;
            _boxes[node.Id] = LayoutEngine.R(box);
        }

        private string Contrast(string fill) =>
            Review.ColorContrast.Ratio(profile.Color("text"), fill) is { } ratio && ratio >= 4.5 ? profile.Color("text") : profile.Color("background");

        /// <summary>Sites: rectangles, diamonds, and flowchart shapes number 1 top, 2 left, 3 bottom, 4 right; 0 asks PowerPoint to reroute.</summary>
        private static int Site(string shape, string side)
        {
            if (side == "center" || shape is "ellipse" or "hexagon" or "database" or "data" or "document")
                return 0;
            return side switch { "top" => 1, "left" => 2, "bottom" => 3, _ => 4 };
        }

        private void AddEdge(DiagramEdge edge, string fromSide, string toSide, string kind = "elbow")
        {
            var from = spec.Nodes.First(node => node.Id == edge.From);
            var to = spec.Nodes.First(node => node.Id == edge.To);
            var key = $"edge-{edge.From}-{edge.To}";
            if (_connectors.Any(connector => connector.Key == key))
                key += $"-{_connectors.Count.ToString(CultureInfo.InvariantCulture)}";
            _connectors.Add(new PlannedElement
            {
                Key = key,
                Role = "diagram-edge",
                Type = "connector",
                Box = _boxes[edge.To],
                FromKey = _nodeKeys[edge.From],
                ToKey = _nodeKeys[edge.To],
                FromSite = Site(from.Shape, fromSide),
                ToSite = Site(to.Shape, toSide),
                ConnectorKind = kind,
                LineColor = profile.Color("muted"),
                LineWidth = 1.25f,
                EndArrow = true,
                Component = Component,
                Tags = new Dictionary<string, string>(Tags("PPTMCP_EDGE", $"{spec.Id}:{edge.From}>{edge.To}"), StringComparer.Ordinal),
            });
            if (edge.Label is { } label)
            {
                var a = _boxes[edge.From];
                var b = _boxes[edge.To];
                var centerX = ((a.Left + a.Right) / 2f + (b.Left + b.Right) / 2f) / 2f;
                var centerY = ((a.Top + a.Bottom) / 2f + (b.Top + b.Bottom) / 2f) / 2f;
                var width = Math.Min(140f, TextEstimator.LineWidth(label, profile.Size("caption")) + 10f);
                var labelKey = $"label-{key}";
                _labels.Add(new PlannedElement
                {
                    Key = labelKey,
                    Role = "diagram-edge-label",
                    Type = "text",
                    Box = LayoutEngine.R(Box.FromSize(centerX - (width / 2f), centerY - (LabelHeight / 2f), width, LabelHeight)),
                    Fill = profile.Color("background"),
                    Paragraphs = [new PlannedParagraph { Text = label, Size = profile.Size("caption"), Font = profile.BodyFont, Color = profile.Color("muted"), Align = "center" }],
                    VAlign = "middle",
                    Padding = [2, 0, 2, 0],
                    Component = Component,
                    Tags = new Dictionary<string, string>(Tags("PPTMCP_EDGE_LABEL", $"{spec.Id}:{edge.From}>{edge.To}"), StringComparer.Ordinal)
                    {
                        [DeckRoles.AllowOverlapTag] = "1",
                    },
                });
                _textKeys.Add(labelKey);
            }
        }

        /// <summary>Layered layout; with lanes, nodes sit in their lane rows and ranks run left to right.</summary>
        public void Layered(Box area, bool horizontal, IReadOnlyList<string>? lanes)
        {
            var ids = spec.Nodes.Select(node => node.Id).ToList();
            var edges = spec.Edges.Select(edge => (edge.From, edge.To)).ToList();
            var (ranks, back) = Rank(ids, edges);
            int rankCount = ranks.Values.Max() + 1;
            var byRank = Enumerable.Range(0, rankCount).Select(rank => ids.Where(id => ranks[id] == rank).ToList()).ToList();

            // Barycenter ordering, two sweeps, ties by original order.
            var order = ids.Select((id, index) => (id, index)).ToDictionary(entry => entry.id, entry => (float)entry.index, StringComparer.Ordinal);
            for (int sweep = 0; sweep < 2; sweep++)
            {
                for (int rank = 1; rank < rankCount; rank++)
                {
                    var current = rank;
                    byRank[current] = byRank[current]
                        .OrderBy(id => Barycenter(id, edges.Where(edge => edge.To == id && ranks[edge.From] == current - 1).Select(edge => edge.From), order))
                        .ThenBy(id => ids.IndexOf(id))
                        .ToList();
                    for (int i = 0; i < byRank[current].Count; i++)
                        order[byRank[current][i]] = i;
                }
            }

            int maxInRank = byRank.Max(rank => rank.Count);
            var gap = profile.Spacing("lg");
            float nodeWidth, nodeHeight;
            var positions = new Dictionary<string, float>(StringComparer.Ordinal);
            if (lanes is not null)
            {
                // Swimlane: x by rank, y by lane (nodes sharing a lane and rank stack within the lane).
                var header = Math.Min(120f, area.Width * 0.16f);
                var content = new Box(area.Left + header, area.Top, area.Right, area.Bottom);
                var laneBoxes = LayoutEngine.Rows(area, lanes.Count, 0);
                var columns = LayoutEngine.Columns(content, rankCount, gap);
                nodeWidth = Math.Min(160f, columns[0].Width - 8f);
                for (int laneIndex = 0; laneIndex < lanes.Count; laneIndex++)
                {
                    var lane = laneBoxes[laneIndex];
                    _backgrounds.Add(new PlannedElement
                    {
                        Key = $"lane-{laneIndex + 1}",
                        Role = "diagram-lane",
                        Type = "shape",
                        Box = lane,
                        Fill = laneIndex % 2 == 0 ? profile.Color("surface") : profile.Color("background"),
                        LineColor = profile.Color("border"),
                        LineWidth = 0.75f,
                        Component = Component,
                        Tags = new Dictionary<string, string>(Tags("PPTMCP_LANE", lanes[laneIndex]), StringComparer.Ordinal),
                    });
                    var headerKey = $"lane-{laneIndex + 1}.header";
                    _backgrounds.Add(new PlannedElement
                    {
                        Key = headerKey,
                        Role = "diagram-lane-header",
                        Type = "shape",
                        Box = LayoutEngine.R(new Box(lane.Left, lane.Top, lane.Left + header, lane.Bottom)),
                        Fill = profile.Color("primary"),
                        Paragraphs = [new PlannedParagraph { Text = lanes[laneIndex], Size = profile.Size("body"), Font = profile.HeadingFont, Color = profile.Color("background"), Bold = true, Align = "center" }],
                        VAlign = "middle",
                        Padding = [4, 2, 4, 2],
                        Component = Component,
                        Tags = new Dictionary<string, string>(Tags("PPTMCP_LANE", lanes[laneIndex]), StringComparer.Ordinal),
                    });
                    _textKeys.Add(headerKey);
                    for (int rank = 0; rank < rankCount; rank++)
                    {
                        var cell = spec.Nodes.Where(node => node.Lane == lanes[laneIndex] && ranks[node.Id] == rank).ToList();
                        if (cell.Count == 0)
                            continue;
                        var slotHeight = (lane.Height - 8f) / cell.Count;
                        nodeHeight = Math.Min(48f, slotHeight - 6f);
                        for (int i = 0; i < cell.Count; i++)
                        {
                            var centerX = (columns[rank].Left + columns[rank].Right) / 2f;
                            var centerY = lane.Top + 4f + (slotHeight * (i + 0.5f));
                            AddNode(cell[i], Box.FromSize(centerX - (nodeWidth / 2f), centerY - (nodeHeight / 2f), nodeWidth, nodeHeight));
                            positions[cell[i].Id] = (laneIndex * 10) + i;
                        }
                    }
                }
                foreach (var edge in spec.Edges)
                {
                    var forward = ranks[edge.To] > ranks[edge.From];
                    var sameRank = ranks[edge.To] == ranks[edge.From];
                    var fromLane = lanes.ToList().IndexOf(spec.Nodes.First(node => node.Id == edge.From).Lane!);
                    var toLane = lanes.ToList().IndexOf(spec.Nodes.First(node => node.Id == edge.To).Lane!);
                    if (forward)
                        AddEdge(edge, "right", "left");
                    else if (sameRank)
                        AddEdge(edge, fromLane < toLane ? "bottom" : "top", fromLane < toLane ? "top" : "bottom");
                    else
                        AddEdge(edge, "bottom", "bottom");
                }
                _notes.Add("Swimlane connectors are PowerPoint elbow connectors glued to node sides; they do not route around other nodes.");
                return;
            }

            var mainSpan = horizontal ? area.Width : area.Height;
            var crossSpan = horizontal ? area.Height : area.Width;
            var crossGap = profile.Spacing("md");
            var crossSize = (crossSpan - (crossGap * (maxInRank - 1))) / maxInRank;
            nodeHeight = horizontal ? Math.Min(56f, crossSize) : Math.Clamp((mainSpan - (gap * (rankCount - 1))) / rankCount, 30f, 56f);
            nodeWidth = horizontal ? Math.Clamp((mainSpan - (gap * (rankCount - 1))) / rankCount, 70f, 170f) : Math.Min(170f, crossSize);
            if (crossSize < (horizontal ? 28f : 70f))
                _notes.Add($"The widest rank has {maxInRank} nodes, so nodes are small ({crossSize:0} pt); consider splitting the diagram.");
            var mainGap = rankCount <= 1 ? 0 : Math.Min(horizontal ? 90f : 70f, (mainSpan - (rankCount * (horizontal ? nodeWidth : nodeHeight))) / (rankCount - 1));
            var mainStart = (horizontal ? area.Left : area.Top) + ((mainSpan - ((rankCount * (horizontal ? nodeWidth : nodeHeight)) + ((rankCount - 1) * mainGap))) / 2f);

            for (int rank = 0; rank < rankCount; rank++)
            {
                var row = byRank[rank];
                var itemSize = horizontal ? nodeHeight : nodeWidth;
                var rowSpan = (row.Count * itemSize) + ((row.Count - 1) * crossGap);
                var crossStart = (horizontal ? area.Top : area.Left) + ((crossSpan - rowSpan) / 2f);
                for (int i = 0; i < row.Count; i++)
                {
                    var node = spec.Nodes.First(item => item.Id == row[i]);
                    var main = mainStart + (rank * ((horizontal ? nodeWidth : nodeHeight) + mainGap));
                    var cross = crossStart + (i * (itemSize + crossGap));
                    var height = node.Shape == "decision" ? nodeHeight * 1.25f : nodeHeight;
                    var box = horizontal
                        ? Box.FromSize(main, cross + ((nodeHeight - height) / 2f), nodeWidth, height)
                        : Box.FromSize(cross, main - ((height - nodeHeight) / 2f), nodeWidth, height);
                    AddNode(node, box);
                    positions[node.Id] = i;
                }
            }

            foreach (var edge in spec.Edges)
            {
                bool isBack = back.Contains((edge.From, edge.To)) || ranks[edge.To] < ranks[edge.From];
                bool sameRank = ranks[edge.To] == ranks[edge.From];
                if (isBack)
                    AddEdge(edge, horizontal ? "bottom" : "left", horizontal ? "bottom" : "left");
                else if (sameRank)
                    AddEdge(edge, horizontal ? "bottom" : "right", horizontal ? "top" : "left");
                else
                    AddEdge(edge, horizontal ? "right" : "bottom", horizontal ? "left" : "top");
                if (!isBack && ranks[edge.To] - ranks[edge.From] > 1)
                    _notes.Add($"Edge {edge.From}→{edge.To} spans {ranks[edge.To] - ranks[edge.From]} ranks and may pass behind nodes in between; PowerPoint's elbow routing does not avoid obstacles.");
            }
            _crossings = CountCrossings(edges.Where(edge => !back.Contains(edge)).ToList(), ranks, positions);
            if (_crossings > 0)
                _notes.Add($"{_crossings} connector crossing(s) remain between adjacent ranks after ordering.");
            if (back.Count > 0)
                _notes.Add($"{back.Count} loop-back edge(s) are drawn around the {(horizontal ? "bottom" : "left")} side.");
        }

        private static float Barycenter(string id, IEnumerable<string> neighbors, Dictionary<string, float> order)
        {
            var list = neighbors.ToList();
            return list.Count == 0 ? order[id] : list.Average(neighbor => order[neighbor]);
        }

        public void Swimlane(Box area) => Layered(area, horizontal: true, lanes: spec.Lanes!);

        public void Matrix(Box area)
        {
            var axisWidth = 28f;
            var grid = LayoutEngine.R(new Box(area.Left + axisWidth, area.Top, area.Right, area.Bottom - axisWidth));
            var columns = LayoutEngine.Columns(grid, 2, 4);
            var rows = LayoutEngine.Rows(grid, 2, 4);
            var cells = new Dictionary<string, Box>(StringComparer.Ordinal)
            {
                ["tl"] = new Box(columns[0].Left, rows[0].Top, columns[0].Right, rows[0].Bottom),
                ["tr"] = new Box(columns[1].Left, rows[0].Top, columns[1].Right, rows[0].Bottom),
                ["bl"] = new Box(columns[0].Left, rows[1].Top, columns[0].Right, rows[1].Bottom),
                ["br"] = new Box(columns[1].Left, rows[1].Top, columns[1].Right, rows[1].Bottom),
            };
            foreach (var (name, cell) in cells)
            {
                var key = $"quadrant-{name}";
                var title = spec.Quadrants?.GetValueOrDefault(name);
                _backgrounds.Add(new PlannedElement
                {
                    Key = key,
                    Role = "diagram-quadrant",
                    Type = "shape",
                    Box = LayoutEngine.R(cell),
                    Fill = name == "tl" ? profile.Color("surface") : profile.Color("background"),
                    LineColor = profile.Color("border"),
                    LineWidth = 1f,
                    Paragraphs = title is null ? null : [new PlannedParagraph { Text = title, Size = profile.Size("caption"), Font = profile.HeadingFont, Color = profile.Color("muted"), Bold = true }],
                    VAlign = "top",
                    Padding = [8, 6, 8, 6],
                    Component = Component,
                    Tags = new Dictionary<string, string>(Tags("PPTMCP_QUADRANT", name), StringComparer.Ordinal),
                });
                if (title is not null)
                    _textKeys.Add(key);
            }
            if (spec.XAxis is { } x)
            {
                _labels.Add(new PlannedElement
                {
                    Key = "axis-x",
                    Role = "diagram-axis",
                    Type = "text",
                    Box = LayoutEngine.R(new Box(grid.Left, grid.Bottom + 4, grid.Right, area.Bottom)),
                    Paragraphs = [new PlannedParagraph { Text = x, Size = profile.Size("body"), Font = profile.HeadingFont, Color = profile.Color("text"), Bold = true, Align = "center" }],
                    Component = Component,
                });
                _textKeys.Add("axis-x");
            }
            if (spec.YAxis is { } y)
            {
                // Rotated 270 degrees around its center: the unrotated box is wide and short.
                var centerX = area.Left + (axisWidth / 2f);
                var centerY = (grid.Top + grid.Bottom) / 2f;
                _labels.Add(new PlannedElement
                {
                    Key = "axis-y",
                    Role = "diagram-axis",
                    Type = "text",
                    Box = LayoutEngine.R(Box.FromSize(centerX - (grid.Height / 2f), centerY - (axisWidth / 2f), grid.Height, axisWidth)),
                    Rotation = 270f,
                    Paragraphs = [new PlannedParagraph { Text = y, Size = profile.Size("body"), Font = profile.HeadingFont, Color = profile.Color("text"), Bold = true, Align = "center" }],
                    VAlign = "middle",
                    Component = Component,
                });
            }

            var chipHeight = 34f;
            foreach (var group in spec.Nodes.Where(node => node.Quadrant is not null).GroupBy(node => node.Quadrant!))
            {
                var cell = LayoutEngine.Inset(cells[group.Key], 10, 28, 10, 8);
                var items = group.ToList();
                int columnsInCell = items.Count > 4 ? 2 : 1;
                var slots = LayoutEngine.Grid(new Box(cell.Left, cell.Top, cell.Right, Math.Min(cell.Bottom, cell.Top + (((items.Count + columnsInCell - 1) / columnsInCell) * (chipHeight + 6f)))), items.Count, columnsInCell, 6, 6);
                for (int i = 0; i < items.Count; i++)
                    AddNode(items[i] with { Shape = items[i].Shape == "process" ? "rounded" : items[i].Shape }, slots[i]);
                if (cell.Top + (((items.Count + columnsInCell - 1) / columnsInCell) * (chipHeight + 6f)) > cell.Bottom)
                    _notes.Add($"Quadrant {group.Key} holds {items.Count} items; chips are compressed.");
            }
            foreach (var node in spec.Nodes.Where(node => node.Quadrant is null))
            {
                var width = Math.Min(150f, grid.Width / 4f);
                var centerX = grid.Left + (node.X!.Value * grid.Width);
                var centerY = grid.Bottom - (node.Y!.Value * grid.Height);
                var box = Box.FromSize(Math.Clamp(centerX - (width / 2f), grid.Left, grid.Right - width), Math.Clamp(centerY - (chipHeight / 2f), grid.Top, grid.Bottom - chipHeight), width, chipHeight);
                AddNode(node with { Shape = node.Shape == "process" ? "rounded" : node.Shape }, box);
            }
        }

        public void HubSpoke(Box area)
        {
            var hub = spec.Nodes.First(node => node.Id == spec.Hub);
            var spokes = spec.Nodes.Where(node => node.Id != spec.Hub).ToList();
            var hubWidth = Math.Min(170f, area.Width * 0.24f);
            var hubHeight = 64f;
            var spokeWidth = Math.Min(150f, area.Width * 0.2f);
            var spokeHeight = 46f;
            var centerX = (area.Left + area.Right) / 2f;
            var centerY = (area.Top + area.Bottom) / 2f;
            AddNode(hub with { Shape = hub.Shape == "process" ? "ellipse" : hub.Shape }, Box.FromSize(centerX - (hubWidth / 2f), centerY - (hubHeight / 2f), hubWidth, hubHeight), profile.Color("primary"));
            var radiusX = (area.Width - spokeWidth) / 2f;
            var radiusY = (area.Height - spokeHeight) / 2f;
            for (int i = 0; i < spokes.Count; i++)
            {
                var angle = (-Math.PI / 2) + (2 * Math.PI * i / spokes.Count);
                var x = centerX + (float)(radiusX * Math.Cos(angle));
                var y = centerY + (float)(radiusY * Math.Sin(angle));
                AddNode(spokes[i] with { Shape = spokes[i].Shape == "process" ? "rounded" : spokes[i].Shape }, Box.FromSize(x - (spokeWidth / 2f), y - (spokeHeight / 2f), spokeWidth, spokeHeight));
                AddEdge(new DiagramEdge(hub.Id, spokes[i].Id), "center", "center", kind: "straight");
            }
            if (spokes.Count > 12)
                _notes.Add($"{spokes.Count} spokes crowd the ring; consider grouping them.");
        }

        public void Architecture(Box area)
        {
            var layers = spec.Layers!;
            var header = Math.Min(120f, area.Width * 0.16f);
            var bands = LayoutEngine.Rows(area, layers.Count, profile.Spacing("sm"));
            for (int i = 0; i < layers.Count; i++)
            {
                var band = bands[i];
                _backgrounds.Add(new PlannedElement
                {
                    Key = $"layer-{i + 1}",
                    Role = "diagram-layer",
                    Type = "shape",
                    Geometry = "rounded-rectangle",
                    Radius = 6,
                    Box = band,
                    Fill = profile.Color("surface"),
                    LineColor = profile.Color("border"),
                    LineWidth = 0.75f,
                    Component = Component,
                    Tags = new Dictionary<string, string>(Tags("PPTMCP_LAYER", layers[i]), StringComparer.Ordinal),
                });
                var labelKey = $"layer-{i + 1}.label";
                _backgrounds.Add(new PlannedElement
                {
                    Key = labelKey,
                    Role = "diagram-layer-label",
                    Type = "text",
                    Box = LayoutEngine.R(new Box(band.Left + 6, band.Top, band.Left + header, band.Bottom)),
                    Paragraphs = [new PlannedParagraph { Text = layers[i], Size = profile.Size("body"), Font = profile.HeadingFont, Color = profile.Color("primary"), Bold = true }],
                    VAlign = "middle",
                    Component = Component,
                });
                _textKeys.Add(labelKey);
                var members = spec.Nodes.Where(node => node.Layer == layers[i]).ToList();
                if (members.Count == 0)
                    continue;
                var inner = LayoutEngine.Inset(new Box(band.Left + header, band.Top, band.Right, band.Bottom), 8, 10, 8, 10);
                var slots = LayoutEngine.Columns(inner, members.Count, profile.Spacing("md"));
                var nodeWidth = Math.Min(170f, slots[0].Width);
                var nodeHeight = Math.Min(54f, inner.Height);
                for (int j = 0; j < members.Count; j++)
                {
                    var slot = slots[j];
                    AddNode(members[j], LayoutEngine.Center(slot, nodeWidth, nodeHeight));
                }
            }
            var layerOf = spec.Nodes.ToDictionary(node => node.Id, node => layers.ToList().IndexOf(node.Layer!), StringComparer.Ordinal);
            foreach (var edge in spec.Edges)
            {
                var down = layerOf[edge.To] > layerOf[edge.From];
                var up = layerOf[edge.To] < layerOf[edge.From];
                if (down)
                    AddEdge(edge, "bottom", "top");
                else if (up)
                    AddEdge(edge, "top", "bottom");
                else
                    AddEdge(edge, "right", "left");
            }
            _notes.Add("Connectors between layers are glued elbow connectors; PowerPoint routes them and they do not avoid other nodes.");
        }
    }
}
