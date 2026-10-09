namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>One planned or applied PPTMCP_ID change.</summary>
public sealed class AppIdAssignment
{
    /// <summary>PowerPoint SlideID of the slide (or of the object's slide).</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based slide position.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>Shape.Id for object assignments; null for slide assignments.</summary>
    public int? ShapeId { get; init; }

    /// <summary>Previous PPTMCP_ID, if any.</summary>
    public string? OldAppId { get; init; }

    /// <summary>New PPTMCP_ID.</summary>
    public required string NewAppId { get; init; }

    /// <summary>missing (no id before) or duplicate (id already used by an earlier slide/object).</summary>
    public required string Reason { get; init; }
}
