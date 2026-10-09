namespace Sbroenne.PowerPointMcp.Core.Data;

/// <summary>A bound table or chart.</summary>
public sealed record BoundObject(string? AppId, int SlideIndex, int SlideId, int ShapeId, string Kind, string Source, bool SourceExists, bool? Stale, string? RefreshedAt);

/// <summary>What refresh changed (or would change) for one object.</summary>
public sealed record RefreshOutcome(string? AppId, int SlideIndex, int ShapeId, string Kind, string Status, int CellsChanged, int RowsAdded, int RowsRemoved, IReadOnlyList<string> Samples, string? Detail);

/// <summary>Result of a data operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class DataOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Columns with inferred types and notes.</summary>
    public IReadOnlyList<DataColumn>? Columns { get; init; }

    /// <summary>Preview rows (display text).</summary>
    public IReadOnlyList<IReadOnlyList<string>>? Rows { get; init; }

    /// <summary>Total data rows in the source (header excluded).</summary>
    public int? TotalRows { get; init; }

    /// <summary>Source hash.</summary>
    public string? Hash { get; init; }

    /// <summary>Slides created (create-table, create-chart).</summary>
    public IReadOnlyList<Composition.ComposedSlide>? Slides { get; init; }

    /// <summary>Bound objects (bindings, bind).</summary>
    public IReadOnlyList<BoundObject>? Bound { get; init; }

    /// <summary>Refresh outcomes.</summary>
    public IReadOnlyList<RefreshOutcome>? Refreshed { get; init; }

    /// <summary>Overflow reported by composition.</summary>
    public IReadOnlyList<string>? Overflow { get; init; }

    /// <summary>Notes on decoding, typing, and fitting.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
