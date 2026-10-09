using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Batch;

/// <summary>Executes one operation through the normal command path (supplied by the service).</summary>
public interface IBatchDispatcher
{
    /// <summary>Runs category.action with camelCase JSON arguments in the batch's session.</summary>
    BatchStepOutcome Dispatch(string category, string action, string? argumentsJson, string jobId);
}

/// <summary>What a dispatched operation returned.</summary>
public sealed record BatchStepOutcome(bool Success, string? ErrorMessage, string? ResultJson);

/// <summary>The outcome of one operation.</summary>
public sealed class BatchStepResult
{
    /// <summary>Operation id.</summary>
    public required string Id { get; init; }

    /// <summary>category.action.</summary>
    public required string Command { get; init; }

    /// <summary>pending, running, ok, failed, skipped, or cancelled.</summary>
    public required string Status { get; set; }

    /// <summary>Error message when failed.</summary>
    public string? Error { get; set; }

    /// <summary>Top-level scalar fields of the result (ids, indexes, counts), or the full result with result_detail=full.</summary>
    public System.Text.Json.Nodes.JsonNode? Result { get; set; }

    /// <summary>Duration in milliseconds.</summary>
    public long? DurationMs { get; set; }
}

/// <summary>State of a batch run, shared between the runner and status queries.</summary>
public sealed class BatchJob
{
    private int _cancelled;

    /// <summary>Job id.</summary>
    public required string Id { get; init; }

    /// <summary>The session object the job runs in.</summary>
    public required object Session { get; init; }

    /// <summary>Per-operation results in order.</summary>
    public required IReadOnlyList<BatchStepResult> Steps { get; init; }

    /// <summary>running, completed, failed, cancelled.</summary>
    public string State { get; set; } = "running";

    /// <summary>When the job started (UTC).</summary>
    public DateTime StartedAt { get; } = DateTime.UtcNow;

    /// <summary>When the job finished (UTC).</summary>
    public DateTime? FinishedAt { get; set; }

    /// <summary>Deck revision before the first operation.</summary>
    public string? RevisionBefore { get; set; }

    /// <summary>Deck revision after the last operation.</summary>
    public string? RevisionAfter { get; set; }

    /// <summary>Checkpoint copy written before the first operation.</summary>
    public string? CheckpointPath { get; set; }

    /// <summary>Structural changes made by the job.</summary>
    public IReadOnlyList<DeckChange>? Changes { get; set; }

    /// <summary>Notes collected while running.</summary>
    public List<string> Notes { get; } = [];

    /// <summary>Whether cancellation was requested (honoured at the next operation boundary).</summary>
    public bool CancellationRequested => Volatile.Read(ref _cancelled) != 0;

    /// <summary>Requests cancellation; the running operation finishes first.</summary>
    public void Cancel() => Interlocked.Exchange(ref _cancelled, 1);

    /// <summary>Operations finished (ok or failed).</summary>
    public int Completed => Steps.Count(step => step.Status is "ok" or "failed");
}

/// <summary>Running and recent batch jobs (kept in memory; the last 50 are retained).</summary>
public static class BatchJobs
{
    private static readonly ConcurrentDictionary<string, BatchJob> Jobs = new(StringComparer.Ordinal);
    private static readonly ConditionalWeakTable<object, BatchJob> Active = new();
    private static readonly Lock Gate = new();

    /// <summary>Registers a job as the session's active job; fails if another is running.</summary>
    public static bool TryStart(BatchJob job, out BatchJob? running)
    {
        ArgumentNullException.ThrowIfNull(job);
        lock (Gate)
        {
            if (Active.TryGetValue(job.Session, out running) && running.State == "running")
                return false;
            Active.AddOrUpdate(job.Session, job);
            Jobs[job.Id] = job;
            foreach (var old in Jobs.Values.Where(item => item.State != "running").OrderByDescending(item => item.StartedAt).Skip(50).ToList())
                Jobs.TryRemove(old.Id, out _);
            running = null;
            return true;
        }
    }

    /// <summary>Marks a job finished and releases the session.</summary>
    public static void Finish(BatchJob job, string state)
    {
        ArgumentNullException.ThrowIfNull(job);
        lock (Gate)
        {
            job.State = state;
            job.FinishedAt = DateTime.UtcNow;
            if (Active.TryGetValue(job.Session, out var current) && ReferenceEquals(current, job))
                Active.Remove(job.Session);
        }
    }

    /// <summary>The job running in a session, if any.</summary>
    public static BatchJob? RunningIn(object session)
    {
        lock (Gate)
        {
            return Active.TryGetValue(session, out var job) && job.State == "running" ? job : null;
        }
    }

    /// <summary>Finds a job by id.</summary>
    public static BatchJob? Find(string id) => Jobs.GetValueOrDefault(id);
}
