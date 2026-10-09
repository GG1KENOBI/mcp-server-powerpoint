namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>
/// One planned repair step. Repairs only move, resize, rescale fonts proportionally, or rewrite
/// ids: they never delete objects, shorten or rewrite text, flatten formatting, or rasterize.
/// </summary>
public sealed class RepairAction
{
    /// <summary>The finding this action addresses.</summary>
    public required string FindingId { get; init; }

    /// <summary>Rule code of that finding.</summary>
    public required string Code { get; init; }

    /// <summary>move, resize, move-resize, scale-font, or assign-id.</summary>
    public required string Operation { get; init; }

    /// <summary>PowerPoint SlideID.</summary>
    public required int SlideId { get; init; }

    /// <summary>1-based slide position when planned.</summary>
    public required int SlideIndex { get; init; }

    /// <summary>Shape.Id of the object to change.</summary>
    public required int ShapeId { get; init; }

    /// <summary>Geometry when planned (left, top, width, height); the repair is skipped if the object moved since.</summary>
    public required IReadOnlyList<float> Before { get; init; }

    /// <summary>Target geometry (left, top, width, height), for move and resize operations.</summary>
    public IReadOnlyList<float>? After { get; init; }

    /// <summary>For scale-font: the smallest font size now.</summary>
    public float? FontSizeBefore { get; init; }

    /// <summary>For scale-font: the estimated smallest font size after scaling. Repair measures and adjusts.</summary>
    public float? FontSizeAfter { get; init; }

    /// <summary>For scale-font: the floor no run may go below.</summary>
    public float? MinFontSize { get; init; }

    /// <summary>For assign-id: the new PPTMCP_ID.</summary>
    public string? NewAppId { get; init; }

    /// <summary>geometry (exact), estimated (verified by measuring during repair), or measured.</summary>
    public required string Basis { get; init; }

    /// <summary>Why this change fixes the finding.</summary>
    public required string Note { get; init; }
}

/// <summary>Outcome of applying one repair action.</summary>
public sealed class RepairOutcome
{
    /// <summary>The action's finding id.</summary>
    public required string FindingId { get; init; }

    /// <summary>The action's operation.</summary>
    public required string Operation { get; init; }

    /// <summary>Shape.Id changed.</summary>
    public required int ShapeId { get; init; }

    /// <summary>applied, skipped, or failed.</summary>
    public required string Status { get; init; }

    /// <summary>What happened, including measured results for font scaling.</summary>
    public required string Detail { get; init; }
}
