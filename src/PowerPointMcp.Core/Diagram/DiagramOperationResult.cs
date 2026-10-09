namespace Sbroenne.PowerPointMcp.Core.Diagram;

/// <summary>A diagram node on the slide.</summary>
public sealed record DiagramNodeInfo(string NodeId, int ShapeId, string AppId, string? Label, IReadOnlyList<float> Box);

/// <summary>A diagram edge on the slide.</summary>
public sealed record DiagramEdgeInfo(string From, string To, int ShapeId, bool Attached, string? Label);

/// <summary>A diagram in the deck.</summary>
public sealed record DiagramSummary(string DiagramId, string Type, int SlideIndex, int SlideId, int Nodes, int Edges);

/// <summary>Result of a diagram operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class DiagramOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Validation errors with JSON paths.</summary>
    public IReadOnlyList<string>? Errors { get; init; }

    /// <summary>Diagram id.</summary>
    public string? DiagramId { get; init; }

    /// <summary>Diagram type.</summary>
    public string? Type { get; init; }

    /// <summary>Slide position.</summary>
    public int? SlideIndex { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public int? SlideId { get; init; }

    /// <summary>Nodes by id with their shapes. Node ids are stable; shape ids change when the diagram is relaid out.</summary>
    public IReadOnlyList<DiagramNodeInfo>? Nodes { get; init; }

    /// <summary>Edges with their connectors and whether both ends are still glued.</summary>
    public IReadOnlyList<DiagramEdgeInfo>? Edges { get; init; }

    /// <summary>Layout and routing notes (crossings, long edges, PowerPoint routing limits).</summary>
    public IReadOnlyList<string>? Notes { get; init; }

    /// <summary>Text that does not fit its node even at the minimum size.</summary>
    public IReadOnlyList<string>? Overflow { get; init; }

    /// <summary>Non-fatal notes.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }

    /// <summary>The stored diagram JSON.</summary>
    public string? Spec { get; init; }

    /// <summary>Diagram types and examples, for types.</summary>
    public IReadOnlyDictionary<string, string>? Examples { get; init; }

    /// <summary>Diagrams in the deck, for list.</summary>
    public IReadOnlyList<DiagramSummary>? Diagrams { get; init; }
}
