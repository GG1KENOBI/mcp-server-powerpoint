using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Attributes;

namespace Sbroenne.PowerPointMcp.Core.Batch;

/// <summary>
/// Ordered batches of tool operations in one session. Every operation is validated before any
/// runs; operations then run one at a time through the normal command path (never two edits at
/// once). Later operations can use earlier results ("$op1.shapeIndex"). Batches are not
/// transactions: a failure stops the batch (or continues, with mode=continue) and nothing is
/// rolled back; use checkpoint=true for a copy you can return to. Batches never save the file.
/// </summary>
[ServiceCategory("batch", "Batch")]
[McpTool("batch", Title = "Batch Operations", Destructive = true, Category = "content",
    Description = "Run an ordered list of tool operations with prevalidation, optional expected_revision check, optional checkpoint copy, per-operation results, partial-failure reporting, and a change log. Use start/status/cancel for long batches with progress. Not a transaction: nothing is rolled back.")]
[McpReadOnlyActions("validate", "status")]
public interface IBatchCommands
{
    /// <summary>Parses and validates operations without running anything.</summary>
    /// <param name="operations">JSON array of {"id","tool","action","args"} objects (args may use snake_case; "$op1.field" uses an earlier result).</param>
    /// <param name="operationsPath">Full path of a local .json file with the operations.</param>
    BatchOperationResult Validate(IPresentationBatch batch, string? operations = null, string? operationsPath = null);

    /// <summary>
    /// Runs the batch and waits for it. Returns per-operation outcomes, the change log, and
    /// revisions. A failure stops the batch unless mode=continue.
    /// </summary>
    /// <param name="expectedRevision">Refuse to start unless the deck revision (deck fingerprint) matches.</param>
    /// <param name="mode">stop (default: stop at the first failure) or continue.</param>
    /// <param name="checkpoint">Save a copy of the presentation before the first operation (default false).</param>
    /// <param name="changeLog">Compare the deck before and after (default true; costs one full read each way).</param>
    /// <param name="resultDetail">summary (default: scalar result fields) or full.</param>
    BatchOperationResult Run(IPresentationBatch batch, string? operations = null, string? operationsPath = null, string? expectedRevision = null,
        string mode = "stop", bool checkpoint = false, bool changeLog = true, string resultDetail = "summary");

    /// <summary>Starts the batch in the background and returns a job id at once. Other edits to this presentation are refused until it finishes.</summary>
    BatchOperationResult Start(IPresentationBatch batch, string? operations = null, string? operationsPath = null, string? expectedRevision = null,
        string mode = "stop", bool checkpoint = false, bool changeLog = true, string resultDetail = "summary");

    /// <summary>Progress and results of a job.</summary>
    /// <param name="jobId">Job id from start or run.</param>
    BatchOperationResult Status(IPresentationBatch batch, string jobId);

    /// <summary>Requests cancellation; the running operation finishes and the rest are marked cancelled.</summary>
    BatchOperationResult Cancel(IPresentationBatch batch, string jobId);
}
