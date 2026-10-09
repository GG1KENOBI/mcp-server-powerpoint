using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Batch;

/// <summary>Batch commands. Constructed per request with the service's dispatcher.</summary>
public sealed class BatchCommands(IBatchDispatcher dispatcher) : IBatchCommands
{
    private const int MaxOperationsBytes = 2 * 1024 * 1024;

    /// <inheritdoc/>
    public BatchOperationResult Validate(IPresentationBatch batch, string? operations = null, string? operationsPath = null)
    {
        var (plan, failure) = Load(operations, operationsPath);
        if (failure is not null)
            return failure;
        return new BatchOperationResult
        {
            Success = true,
            Total = plan!.Operations.Count,
            Operations = plan.Operations.Select(operation => new BatchStepResult { Id = operation.Id, Command = operation.Command, Status = "valid" }).ToList(),
            Warnings = plan.Warnings.Count > 0 ? plan.Warnings : null,
        };
    }

    /// <inheritdoc/>
    public BatchOperationResult Run(IPresentationBatch batch, string? operations = null, string? operationsPath = null, string? expectedRevision = null,
        string mode = "stop", bool checkpoint = false, bool changeLog = true, string resultDetail = "summary")
    {
        var (job, plan, failure) = Begin(batch, operations, operationsPath, mode, resultDetail);
        if (failure is not null)
            return failure;
        Execute(batch, job!, plan!, expectedRevision, mode, checkpoint, changeLog, resultDetail);
        return Report(job!, plan!.Warnings);
    }

    /// <inheritdoc/>
    public BatchOperationResult Start(IPresentationBatch batch, string? operations = null, string? operationsPath = null, string? expectedRevision = null,
        string mode = "stop", bool checkpoint = false, bool changeLog = true, string resultDetail = "summary")
    {
        var (job, plan, failure) = Begin(batch, operations, operationsPath, mode, resultDetail);
        if (failure is not null)
            return failure;
        _ = Task.Run(() => Execute(batch, job!, plan!, expectedRevision, mode, checkpoint, changeLog, resultDetail));
        return new BatchOperationResult
        {
            Success = true,
            JobId = job!.Id,
            State = "running",
            Total = plan!.Operations.Count,
            Completed = 0,
            Warnings = [$"Started. Poll batch status with job_id {job.Id}; edits to this presentation from other calls are refused until the job finishes."],
        };
    }

    /// <inheritdoc/>
    public BatchOperationResult Status(IPresentationBatch batch, string jobId)
    {
        var job = BatchJobs.Find(jobId);
        return job is null ? Fail($"No batch job has id '{jobId}' (jobs are kept in memory until the server restarts; the last 50 finished jobs are retained).") : Report(job, null);
    }

    /// <inheritdoc/>
    public BatchOperationResult Cancel(IPresentationBatch batch, string jobId)
    {
        var job = BatchJobs.Find(jobId);
        if (job is null)
            return Fail($"No batch job has id '{jobId}'.");
        if (job.State != "running")
            return Fail($"Job {jobId} already finished ({job.State}).");
        job.Cancel();
        return new BatchOperationResult
        {
            Success = true,
            JobId = jobId,
            State = "cancelling",
            Completed = job.Completed,
            Total = job.Steps.Count,
            Warnings = ["The running operation finishes first (operations are never interrupted half-way); the remaining operations are marked cancelled. Completed operations are not undone."],
        };
    }

    private static (BatchJob? Job, BatchPlan? Plan, BatchOperationResult? Failure) Begin(IPresentationBatch batch, string? operations, string? operationsPath, string mode, string resultDetail)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (mode is not ("stop" or "continue"))
            return (null, null, Fail("mode must be stop or continue."));
        if (resultDetail is not ("summary" or "full"))
            return (null, null, Fail("result_detail must be summary or full."));
        var (plan, failure) = Load(operations, operationsPath);
        if (failure is not null)
            return (null, null, failure);
        var job = new BatchJob
        {
            Id = "job-" + Guid.NewGuid().ToString("N")[..10],
            Session = batch,
            Steps = plan!.Operations.Select(operation => new BatchStepResult { Id = operation.Id, Command = operation.Command, Status = "pending" }).ToList(),
        };
        if (!BatchJobs.TryStart(job, out var running))
            return (null, null, Fail($"Batch job {running!.Id} is still running in this session ({running.Completed}/{running.Steps.Count} done). Wait for it (batch status) or cancel it."));
        return (job, plan, null);
    }

    private void Execute(IPresentationBatch batch, BatchJob job, BatchPlan plan, string? expectedRevision, string mode, bool checkpoint, bool changeLog, string resultDetail)
    {
        string final = "failed";
        bool reachedEnd = false;
        try
        {
            List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)>? before = null;
            if (changeLog || expectedRevision is not null)
            {
                before = batch.Execute((ctx, ct) => DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: false, ct));
                job.RevisionBefore = DeckCommands.Revision(before);
                if (expectedRevision is not null && expectedRevision != job.RevisionBefore)
                {
                    foreach (var step in job.Steps)
                        step.Status = "skipped";
                    job.Notes.Add($"Not started: the deck changed since revision {expectedRevision} (now {job.RevisionBefore}).");
                    return;
                }
            }

            if (checkpoint)
            {
                var folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerPointMcp", "checkpoints");
                Directory.CreateDirectory(folder);
                job.CheckpointPath = batch.Execute((ctx, ct) =>
                {
                    var name = Path.GetFileNameWithoutExtension(ctx.Presentation.Name);
                    var extension = Path.GetExtension(ctx.Presentation.Name) is { Length: > 0 } ext ? ext : ".pptx";
                    var path = Path.Combine(folder, $"{name}-checkpoint-{DateTime.Now:yyyyMMdd-HHmmss}-{job.Id}{extension}");
                    // SaveCopyAs writes the in-memory state (including unsaved edits) without changing the open file's path.
                    ctx.Presentation.SaveCopyAs(path);
                    return path;
                });
            }

            var results = new Dictionary<string, JsonNode?>(StringComparer.Ordinal);
            bool stopped = false;
            for (int index = 0; index < plan.Operations.Count; index++)
            {
                var operation = plan.Operations[index];
                var step = job.Steps[index];
                if (stopped)
                {
                    step.Status = "skipped";
                    continue;
                }
                if (job.CancellationRequested)
                {
                    step.Status = "cancelled";
                    continue;
                }

                step.Status = "running";
                var watch = Stopwatch.StartNew();
                var (arguments, referenceError) = BatchPlanner.Resolve(operation, results);
                if (referenceError is not null)
                {
                    step.Status = "failed";
                    step.Error = referenceError;
                }
                else
                {
                    var outcome = dispatcher.Dispatch(operation.Category, operation.Action, arguments, job.Id);
                    var node = Parse(outcome.ResultJson);
                    results[operation.Id] = node;
                    step.Status = outcome.Success ? "ok" : "failed";
                    step.Error = outcome.Success ? null : outcome.ErrorMessage ?? "Operation failed.";
                    step.Result = resultDetail == "full" ? node : Summary(node);
                }
                step.DurationMs = watch.ElapsedMilliseconds;
                if (step.Status == "failed" && mode == "stop")
                    stopped = true;
            }

            if (before is not null && changeLog)
            {
                var after = batch.Execute((ctx, ct) => DeckSnapshotReader.ReadAll(ctx.Presentation, detailed: false, ct));
                job.RevisionAfter = DeckCommands.Revision(after);
                job.Changes = DeckDiff.Compare(before, after);
            }

            final = job.Steps.Any(step => step.Status == "cancelled") ? "cancelled"
                : job.Steps.Any(step => step.Status is "failed" or "skipped") ? "failed"
                : "completed";
            reachedEnd = true;
        }
        finally
        {
            if (!reachedEnd)
            {
                foreach (var step in job.Steps.Where(step => step.Status is "running" or "pending"))
                    step.Status = step.Status == "running" ? "failed" : "skipped";
                job.Notes.Add("The batch stopped because of an unexpected error outside the operations (reading the deck or writing the checkpoint); see the server log. Completed operations were not undone.");
            }
            BatchJobs.Finish(job, final);
        }
    }

    private static (BatchPlan? Plan, BatchOperationResult? Failure) Load(string? operations, string? operationsPath)
    {
        if (operations is not null && operationsPath is not null)
            return (null, Fail("Pass operations or operations_path, not both."));
        string json;
        if (operationsPath is not null)
        {
            if (!File.Exists(operationsPath))
                return (null, Fail($"operations_path '{operationsPath}' does not exist."));
            if (new FileInfo(operationsPath).Length > MaxOperationsBytes)
                return (null, Fail("operations_path is larger than 2 MB; split the batch."));
            json = File.ReadAllText(operationsPath);
        }
        else if (operations is not null)
        {
            json = operations;
        }
        else
        {
            return (null, Fail("Pass the operations JSON array in operations, or a file path in operations_path."));
        }

        var plan = BatchPlanner.Parse(json);
        if (plan.Errors.Count > 0)
        {
            return (null, new BatchOperationResult
            {
                Success = false,
                ErrorMessage = $"The batch is invalid ({plan.Errors.Count} problem(s)); nothing ran. First: {plan.Errors[0]}",
                Errors = plan.Errors,
                Warnings = plan.Warnings.Count > 0 ? plan.Warnings : null,
            });
        }
        return (plan, null);
    }

    private static BatchOperationResult Report(BatchJob job, IReadOnlyList<string>? warnings)
    {
        int failed = job.Steps.Count(step => step.Status == "failed");
        bool finished = job.State != "running";
        bool ok = finished && job.State == "completed";
        var notes = new List<string>(warnings ?? []);
        notes.AddRange(job.Notes);
        string? error = null;
        if (finished && !ok)
        {
            var firstFailure = job.Steps.FirstOrDefault(step => step.Status == "failed");
            error = job.State switch
            {
                "cancelled" => $"Cancelled after {job.Completed} of {job.Steps.Count} operations; completed operations were not undone.",
                _ when firstFailure is not null => $"{failed} of {job.Steps.Count} operations failed; first: {firstFailure.Id} ({firstFailure.Command}): {firstFailure.Error}. Completed operations were not undone.",
                _ => job.Notes.FirstOrDefault() ?? "The batch did not run.",
            };
        }
        return new BatchOperationResult
        {
            Success = !finished || ok,
            ErrorMessage = error,
            JobId = job.Id,
            State = job.State,
            Total = job.Steps.Count,
            Completed = job.Completed,
            Failed = failed,
            Operations = job.Steps,
            RevisionBefore = job.RevisionBefore,
            RevisionAfter = job.RevisionAfter,
            Changes = job.Changes,
            CheckpointPath = job.CheckpointPath,
            RestoreSteps = job.CheckpointPath is null || ok ? null :
            [
                $"The checkpoint {job.CheckpointPath} holds the presentation as it was before the batch (including unsaved edits).",
                "To return to it: presentation close with save=false (this discards the batch's changes in the open file), then presentation open with file_path set to the checkpoint.",
                "Then presentation save-as to the original path if you want the original file to match the checkpoint. Nothing is restored automatically.",
            ],
            Warnings = notes.Count > 0 ? notes : null,
        };
    }

    private static JsonNode? Parse(string? json)
    {
        if (string.IsNullOrWhiteSpace(json))
            return null;
        try
        {
            return JsonNode.Parse(json);
        }
        catch (JsonException)
        {
            return JsonValue.Create(json);
        }
    }

    /// <summary>Keeps scalar top-level fields (ids, indexes, counts, paths) and drops large arrays and objects.</summary>
    internal static JsonNode? Summary(JsonNode? node)
    {
        if (node is not JsonObject obj)
            return node;
        var summary = new JsonObject();
        foreach (var (name, value) in obj)
        {
            if (value is JsonValue scalar && name is not "success" and not "errorMessage")
                summary[name] = scalar.DeepClone();
        }
        return summary;
    }

    private static BatchOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
