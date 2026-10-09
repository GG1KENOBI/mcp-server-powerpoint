using System.Text.Json;
using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Diagram;

/// <summary>A diagram node. The id is the stable address used by update-node and stored in PPTMCP_NODE.</summary>
public sealed record DiagramNode(
    string Id,
    string Label,
    string? Detail = null,
    string Shape = "process",
    string? Lane = null,
    string? Layer = null,
    string? Quadrant = null,
    float? X = null,
    float? Y = null,
    string? Tone = null);

/// <summary>A directed edge between two node ids.</summary>
public sealed record DiagramEdge(string From, string To, string? Label = null, bool Dashed = false);

/// <summary>A validated diagram (schema pptmcp.diagram/1).</summary>
public sealed record DiagramSpec
{
    /// <summary>Schema identifier.</summary>
    public const string SchemaId = "pptmcp.diagram/1";

    /// <summary>flowchart, swimlane, matrix, hub-spoke, or architecture.</summary>
    public required string Type { get; init; }

    /// <summary>Diagram id (unique in the deck).</summary>
    public required string Id { get; init; }

    /// <summary>Title for a new slide.</summary>
    public string? Title { get; init; }

    /// <summary>Flowchart direction: down (default) or right.</summary>
    public string Direction { get; init; } = "down";

    /// <summary>Nodes in order.</summary>
    public required IReadOnlyList<DiagramNode> Nodes { get; init; }

    /// <summary>Edges in order.</summary>
    public IReadOnlyList<DiagramEdge> Edges { get; init; } = [];

    /// <summary>Swimlane lanes, top to bottom.</summary>
    public IReadOnlyList<string>? Lanes { get; init; }

    /// <summary>Architecture layers, top to bottom.</summary>
    public IReadOnlyList<string>? Layers { get; init; }

    /// <summary>Matrix axis titles.</summary>
    public string? XAxis { get; init; }

    /// <summary>Matrix axis titles.</summary>
    public string? YAxis { get; init; }

    /// <summary>Matrix quadrant titles by tl, tr, bl, br.</summary>
    public IReadOnlyDictionary<string, string>? Quadrants { get; init; }

    /// <summary>Hub node id (hub-spoke).</summary>
    public string? Hub { get; init; }

    /// <summary>Area on the slide in points; default: the slide body below the title.</summary>
    public Box? Area { get; init; }

    /// <summary>Design profile name.</summary>
    public string? Profile { get; init; }
}

/// <summary>Parses and validates diagram JSON.</summary>
public static partial class DiagramParser
{
    /// <summary>Diagram types.</summary>
    public static readonly string[] Types = ["flowchart", "swimlane", "matrix", "hub-spoke", "architecture"];

    /// <summary>Node shapes.</summary>
    public static readonly string[] Shapes = ["process", "rounded", "decision", "terminator", "data", "document", "database", "ellipse", "hexagon"];

    private static readonly string[] Directions = ["down", "right"];
    private static readonly string[] QuadrantNames = ["tl", "tr", "bl", "br"];
    private static readonly string[] Tones = ["primary", "secondary", "accent", "positive", "negative", "neutral", "surface"];

    /// <summary>Examples by type.</summary>
    public static readonly IReadOnlyDictionary<string, string> Examples = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["flowchart"] = """{"type":"flowchart","id":"approval","title":"Approval flow","nodes":[{"id":"start","label":"Request","shape":"terminator"},{"id":"check","label":"Budget < 10k?","shape":"decision"},{"id":"auto","label":"Auto-approve"},{"id":"cfo","label":"CFO review"},{"id":"done","label":"Done","shape":"terminator"}],"edges":[{"from":"start","to":"check"},{"from":"check","to":"auto","label":"yes"},{"from":"check","to":"cfo","label":"no"},{"from":"auto","to":"done"},{"from":"cfo","to":"done"}]}""",
        ["swimlane"] = """{"type":"swimlane","id":"order","title":"Order to cash","lanes":["Customer","Sales","Finance"],"nodes":[{"id":"po","label":"Place order","lane":"Customer"},{"id":"confirm","label":"Confirm order","lane":"Sales"},{"id":"invoice","label":"Send invoice","lane":"Finance"},{"id":"pay","label":"Pay","lane":"Customer"}],"edges":[{"from":"po","to":"confirm"},{"from":"confirm","to":"invoice"},{"from":"invoice","to":"pay"}]}""",
        ["matrix"] = """{"type":"matrix","id":"prio","title":"Prioritisation","x_axis":"Effort","y_axis":"Impact","quadrants":{"tl":"Quick wins","tr":"Big bets","bl":"Fill-ins","br":"Avoid"},"nodes":[{"id":"a","label":"Self-service portal","quadrant":"tl"},{"id":"b","label":"New ERP","quadrant":"tr"},{"id":"c","label":"Logo refresh","quadrant":"bl"}]}""",
        ["hub-spoke"] = """{"type":"hub-spoke","id":"eco","title":"Partner ecosystem","hub":"core","nodes":[{"id":"core","label":"Platform"},{"id":"p1","label":"Banks"},{"id":"p2","label":"Retail"},{"id":"p3","label":"Logistics"},{"id":"p4","label":"Government"}]}""",
        ["architecture"] = """{"type":"architecture","id":"arch","title":"Solution architecture","layers":["Channels","Services","Data"],"nodes":[{"id":"web","label":"Web app","layer":"Channels"},{"id":"mobile","label":"Mobile app","layer":"Channels"},{"id":"api","label":"API gateway","layer":"Services"},{"id":"orders","label":"Order service","layer":"Services"},{"id":"db","label":"PostgreSQL","layer":"Data","shape":"database"}],"edges":[{"from":"web","to":"api"},{"from":"mobile","to":"api"},{"from":"api","to":"orders"},{"from":"orders","to":"db"}]}""",
    };

    /// <summary>Parses diagram JSON; errors carry JSON paths.</summary>
    public static (DiagramSpec? Spec, IReadOnlyList<string> Errors, IReadOnlyList<string> Warnings) Parse(string json)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        JsonDocument document;
        try
        {
            document = JsonDocument.Parse(json ?? "", new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip, AllowTrailingCommas = true });
        }
        catch (JsonException ex)
        {
            return (null, [$"$: invalid JSON at line {(ex.LineNumber ?? 0) + 1}: {ex.Message}"], []);
        }

        using (document)
        {
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object)
                return (null, ["$: the diagram must be a JSON object."], []);
            var reader = new JsonSpecReader(errors, warnings);
            reader.Unknown(root, "$", "schema", "type", "id", "title", "direction", "nodes", "edges", "lanes", "layers", "x_axis", "y_axis", "quadrants", "hub", "area", "profile");
            var schema = reader.String(root, "$", "schema", required: false);
            if (schema is not null && schema != DiagramSpec.SchemaId)
                errors.Add($"$.schema: expected {DiagramSpec.SchemaId}.");
            var type = reader.Enum(root, "$", "type", Types);
            if (type is null && !errors.Any(error => error.StartsWith("$.type", StringComparison.Ordinal)))
                errors.Add($"$.type: required, one of {string.Join(", ", Types)}.");
            var id = reader.String(root, "$", "id", required: true);
            if (id is not null && !DiagramId().IsMatch(id))
                errors.Add("$.id: use 1-40 letters, digits, dash, or underscore.");

            var nodes = reader.Array(root, "$", "nodes", (element, path) =>
            {
                if (!reader.IsObject(element, path))
                    return null;
                reader.Unknown(element, path, "id", "label", "detail", "shape", "lane", "layer", "quadrant", "x", "y", "tone");
                var nodeId = reader.String(element, path, "id", required: true);
                var label = reader.String(element, path, "label", required: true);
                if (nodeId is not null && !NodeId().IsMatch(nodeId))
                    errors.Add($"{path}.id: use 1-40 letters, digits, dot, dash, or underscore.");
                return nodeId is null || label is null
                    ? null
                    : new DiagramNode(nodeId, label,
                        reader.String(element, path, "detail", required: false),
                        reader.Enum(element, path, "shape", Shapes) ?? "process",
                        reader.String(element, path, "lane", required: false),
                        reader.String(element, path, "layer", required: false),
                        reader.Enum(element, path, "quadrant", QuadrantNames),
                        reader.Float(element, path, "x", 0f, 1f),
                        reader.Float(element, path, "y", 0f, 1f),
                        reader.Enum(element, path, "tone", Tones));
            }) ?? [];
            if (nodes.Count == 0)
                errors.Add("$.nodes: at least one node is required.");
            if (nodes.Count > 40)
                errors.Add($"$.nodes: {nodes.Count} nodes do not fit legibly on one slide (max 40); split the diagram.");
            foreach (var duplicate in nodes.GroupBy(node => node.Id).Where(group => group.Count() > 1))
                errors.Add($"$.nodes: id '{duplicate.Key}' is used {duplicate.Count()} times.");

            var ids = nodes.Select(node => node.Id).ToHashSet(StringComparer.Ordinal);
            var edges = reader.Array(root, "$", "edges", (element, path) =>
            {
                if (!reader.IsObject(element, path))
                    return null;
                reader.Unknown(element, path, "from", "to", "label", "dashed");
                var from = reader.String(element, path, "from", required: true);
                var to = reader.String(element, path, "to", required: true);
                if (from is not null && !ids.Contains(from))
                    errors.Add($"{path}.from: no node has id '{from}'.");
                if (to is not null && !ids.Contains(to))
                    errors.Add($"{path}.to: no node has id '{to}'.");
                if (from is not null && from == to)
                    errors.Add($"{path}: an edge from a node to itself cannot be drawn with a glued connector.");
                return from is null || to is null ? null : new DiagramEdge(from, to, reader.String(element, path, "label", required: false), reader.Bool(element, path, "dashed") ?? false);
            }) ?? [];

            var lanes = reader.StringArray(root, "$", "lanes");
            var layers = reader.StringArray(root, "$", "layers");
            var hub = reader.String(root, "$", "hub", required: false);
            Box? area = null;
            if (root.TryGetProperty("area", out var areaElement) && areaElement.ValueKind == JsonValueKind.Object)
            {
                var left = reader.Float(areaElement, "$.area", "left", 0, 5000);
                var top = reader.Float(areaElement, "$.area", "top", 0, 5000);
                var width = reader.Float(areaElement, "$.area", "width", 50, 5000);
                var height = reader.Float(areaElement, "$.area", "height", 50, 5000);
                if (left is null || top is null || width is null || height is null)
                    errors.Add("$.area: needs left, top, width, and height in points.");
                else
                    area = Box.FromSize(left.Value, top.Value, width.Value, height.Value);
            }
            Dictionary<string, string>? quadrants = null;
            if (root.TryGetProperty("quadrants", out var quadrantElement) && quadrantElement.ValueKind == JsonValueKind.Object)
            {
                quadrants = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var property in quadrantElement.EnumerateObject())
                {
                    if (Array.IndexOf(QuadrantNames, property.Name) < 0 || property.Value.ValueKind != JsonValueKind.String)
                        errors.Add($"$.quadrants.{property.Name}: keys are tl, tr, bl, br with string values.");
                    else
                        quadrants[property.Name] = property.Value.GetString()!;
                }
            }

            switch (type)
            {
                case "swimlane":
                    var laneNames = lanes ?? nodes.Select(node => node.Lane).OfType<string>().Distinct().ToList();
                    if (laneNames.Count == 0)
                        errors.Add("$.lanes: a swimlane needs lanes (or a lane on every node).");
                    foreach (var node in nodes.Where(node => node.Lane is null || !laneNames.Contains(node.Lane)))
                        errors.Add($"$.nodes[{IndexOf(nodes, node)}].lane: '{node.Lane}' is not one of the lanes ({string.Join(", ", laneNames)}).");
                    lanes = laneNames;
                    break;
                case "architecture":
                    var layerNames = layers ?? nodes.Select(node => node.Layer).OfType<string>().Distinct().ToList();
                    if (layerNames.Count == 0)
                        errors.Add("$.layers: an architecture diagram needs layers (or a layer on every node).");
                    foreach (var node in nodes.Where(node => node.Layer is null || !layerNames.Contains(node.Layer)))
                        errors.Add($"$.nodes[{IndexOf(nodes, node)}].layer: '{node.Layer}' is not one of the layers ({string.Join(", ", layerNames)}).");
                    layers = layerNames;
                    break;
                case "matrix":
                    foreach (var node in nodes.Where(node => node.Quadrant is null && (node.X is null || node.Y is null)))
                        errors.Add($"$.nodes[{IndexOf(nodes, node)}]: matrix nodes need a quadrant (tl, tr, bl, br) or x and y between 0 and 1.");
                    if (edges.Count > 0)
                        warnings.Add("$.edges: matrix diagrams do not draw edges; ignored.");
                    break;
                case "hub-spoke":
                    if (hub is null || !ids.Contains(hub))
                        errors.Add($"$.hub: required, the id of the center node.");
                    if (edges.Count > 0)
                        warnings.Add("$.edges: hub-spoke connects every node to the hub; edges are ignored.");
                    break;
            }

            var direction = reader.Enum(root, "$", "direction", Directions) ?? (type == "swimlane" ? "right" : "down");
            if (errors.Count > 0)
                return (null, errors, warnings);
            return (new DiagramSpec
            {
                Type = type!,
                Id = id!,
                Title = reader.String(root, "$", "title", required: false),
                Direction = direction,
                Nodes = nodes,
                Edges = type is "matrix" or "hub-spoke" ? [] : edges,
                Lanes = lanes,
                Layers = layers,
                XAxis = reader.String(root, "$", "x_axis", required: false),
                YAxis = reader.String(root, "$", "y_axis", required: false),
                Quadrants = quadrants,
                Hub = hub,
                Area = area,
                Profile = reader.String(root, "$", "profile", required: false),
            }, errors, warnings);
        }
    }

    /// <summary>Serializes a spec back to compact JSON (stored on the slide).</summary>
    public static string ToJson(DiagramSpec spec)
    {
        ArgumentNullException.ThrowIfNull(spec);
        var payload = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["schema"] = DiagramSpec.SchemaId,
            ["type"] = spec.Type,
            ["id"] = spec.Id,
            ["title"] = spec.Title,
            ["direction"] = spec.Direction,
            ["nodes"] = spec.Nodes.Select(node => new Dictionary<string, object?>
            {
                ["id"] = node.Id,
                ["label"] = node.Label,
                ["detail"] = node.Detail,
                ["shape"] = node.Shape,
                ["lane"] = node.Lane,
                ["layer"] = node.Layer,
                ["quadrant"] = node.Quadrant,
                ["x"] = node.X,
                ["y"] = node.Y,
                ["tone"] = node.Tone,
            }.Where(entry => entry.Value is not null).ToDictionary(entry => entry.Key, entry => entry.Value)),
            ["edges"] = spec.Edges.Select(edge => new Dictionary<string, object?>
            {
                ["from"] = edge.From,
                ["to"] = edge.To,
                ["label"] = edge.Label,
                ["dashed"] = edge.Dashed ? true : null,
            }.Where(entry => entry.Value is not null).ToDictionary(entry => entry.Key, entry => entry.Value)),
            ["lanes"] = spec.Lanes,
            ["layers"] = spec.Layers,
            ["x_axis"] = spec.XAxis,
            ["y_axis"] = spec.YAxis,
            ["quadrants"] = spec.Quadrants,
            ["hub"] = spec.Hub,
            ["area"] = spec.Area is { } area ? new { left = area.Left, top = area.Top, width = area.Width, height = area.Height } : null,
            ["profile"] = spec.Profile,
        };
        return JsonSerializer.Serialize(payload.Where(entry => entry.Value is not null).ToDictionary(entry => entry.Key, entry => entry.Value));
    }

    private static int IndexOf(List<DiagramNode> nodes, DiagramNode node)
    {
        for (int i = 0; i < nodes.Count; i++)
        {
            if (ReferenceEquals(nodes[i], node))
                return i;
        }
        return -1;
    }

    [GeneratedRegex("^[A-Za-z0-9_-]{1,40}$", RegexOptions.CultureInvariant)]
    private static partial Regex DiagramId();

    [GeneratedRegex("^[A-Za-z0-9_.-]{1,40}$", RegexOptions.CultureInvariant)]
    private static partial Regex NodeId();
}
