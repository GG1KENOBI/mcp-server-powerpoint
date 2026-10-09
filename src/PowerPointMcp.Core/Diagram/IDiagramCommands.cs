using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Diagram;

/// <summary>
/// Editable diagrams (schema pptmcp.diagram/1): flowcharts with branches and loops, swimlanes,
/// 2x2 matrices, hub-and-spoke, and layered architecture. Nodes are native shapes tagged with
/// stable node ids; edges are connectors glued to node connection sites so they follow moved
/// nodes. Update nodes by id; structural edits relayout deterministically.
/// </summary>
[ServiceCategory("diagram", "Diagram")]
[McpTool("diagram", Title = "Diagrams", Destructive = true, Category = "content",
    Description = "Create flowchart, swimlane, matrix, hub-spoke, and architecture diagrams from JSON with glued connectors and stable node ids; inspect them; change node text by id; add or remove nodes and edges; relayout. Use types for examples.")]
[McpReadOnlyActions("types", "inspect", "list")]
public interface IDiagramCommands
{
    /// <summary>Lists diagram types with a complete example for each.</summary>
    DiagramOperationResult Types(IPresentationBatch batch);

    /// <summary>
    /// Creates a diagram on a slide (slide_index or slide_id), or on a new title-only slide after
    /// the last one when neither is given. The area defaults to the slide body below the title.
    /// </summary>
    /// <param name="spec">Diagram JSON (pptmcp.diagram/1). Alternatively pass spec_path.</param>
    /// <param name="specPath">Full path of a local .json file holding the diagram.</param>
    /// <param name="slideIndex">1-based slide to draw on.</param>
    /// <param name="slideId">PowerPoint SlideID to draw on.</param>
    /// <param name="profile">Design profile (default theme).</param>
    DiagramOperationResult Create(IPresentationBatch batch, string? spec = null, string? specPath = null, int? slideIndex = null, int? slideId = null, string? profile = null);

    /// <summary>Returns the diagram's nodes, edges (with glue state), and stored JSON.</summary>
    /// <param name="diagramId">Diagram id.</param>
    DiagramOperationResult Inspect(IPresentationBatch batch, string diagramId);

    /// <summary>Lists diagrams in the deck.</summary>
    DiagramOperationResult List(IPresentationBatch batch);

    /// <summary>
    /// Changes a node's label and/or detail in place (formatting kept). A new tone or shape
    /// relayouts the diagram.
    /// </summary>
    /// <param name="nodeId">Node id.</param>
    /// <param name="label">New label.</param>
    /// <param name="detail">New detail line (empty string removes it).</param>
    /// <param name="tone">Node color token: primary, secondary, accent, positive, negative, neutral, surface.</param>
    /// <param name="shape">process, rounded, decision, terminator, data, document, database, ellipse, hexagon.</param>
    DiagramOperationResult UpdateNode(IPresentationBatch batch, string diagramId, string nodeId, string? label = null, [AllowEmptyString] string? detail = null, string? tone = null, string? shape = null);

    /// <summary>Adds a node (optionally connected from/to existing nodes) and relayouts.</summary>
    /// <param name="lane">Lane for swimlanes.</param>
    /// <param name="layer">Layer for architecture diagrams.</param>
    /// <param name="quadrant">tl, tr, bl, or br for matrices.</param>
    /// <param name="connectFrom">Existing node id to draw an edge from.</param>
    /// <param name="connectTo">Existing node id to draw an edge to.</param>
    DiagramOperationResult AddNode(IPresentationBatch batch, string diagramId, string nodeId, string label, string? detail = null, string? shape = null,
        string? lane = null, string? layer = null, string? quadrant = null, string? connectFrom = null, string? connectTo = null);

    /// <summary>Removes a node and its edges, then relayouts.</summary>
    DiagramOperationResult RemoveNode(IPresentationBatch batch, string diagramId, string nodeId);

    /// <summary>Adds an edge and relayouts.</summary>
    /// <param name="fromNode">Start node id.</param>
    /// <param name="toNode">End node id.</param>
    /// <param name="edgeLabel">Label shown on the edge.</param>
    DiagramOperationResult AddEdge(IPresentationBatch batch, string diagramId, string fromNode, string toNode, string? edgeLabel = null);

    /// <summary>Removes the edge(s) between two nodes and relayouts.</summary>
    DiagramOperationResult RemoveEdge(IPresentationBatch batch, string diagramId, string fromNode, string toNode);

    /// <summary>Recreates the diagram from its stored JSON in its stored area (manual formatting of its shapes is not kept).</summary>
    DiagramOperationResult Relayout(IPresentationBatch batch, string diagramId, string? profile = null);
}
