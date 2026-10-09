using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Batch;

/// <summary>Result of a batch operation. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class BatchOperationResult
{
    /// <summary>Whether every operation succeeded (or, for dry runs and starts, whether validation passed).</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false: validation problems or which operations failed.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Validation problems (operation id: message); nothing ran when these are present.</summary>
    public IReadOnlyList<string>? Errors { get; init; }

    /// <summary>Job id (status and cancel use it).</summary>
    public string? JobId { get; init; }

    /// <summary>running, completed, failed, or cancelled.</summary>
    public string? State { get; init; }

    /// <summary>Operations total.</summary>
    public int? Total { get; init; }

    /// <summary>Operations finished (ok or failed).</summary>
    public int? Completed { get; init; }

    /// <summary>Operations that failed.</summary>
    public int? Failed { get; init; }

    /// <summary>Per-operation outcomes.</summary>
    public IReadOnlyList<BatchStepResult>? Operations { get; init; }

    /// <summary>Deck revision before the batch.</summary>
    public string? RevisionBefore { get; init; }

    /// <summary>Deck revision after the batch.</summary>
    public string? RevisionAfter { get; init; }

    /// <summary>Structural changes (slides and objects added, removed, moved, modified).</summary>
    public IReadOnlyList<DeckChange>? Changes { get; init; }

    /// <summary>Checkpoint copy written before the first operation.</summary>
    public string? CheckpointPath { get; init; }

    /// <summary>How to return to the checkpoint. Batches are not transactions: nothing is rolled back automatically.</summary>
    public IReadOnlyList<string>? RestoreSteps { get; init; }

    /// <summary>Non-fatal notes.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
