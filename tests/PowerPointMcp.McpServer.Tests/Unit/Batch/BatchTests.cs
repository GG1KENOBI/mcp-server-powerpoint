// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Text.Json.Nodes;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Batch;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Batch;

/// <summary>Batch parsing, prevalidation, references, failure modes, and jobs with a fake dispatcher (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Batch")]
public sealed class BatchTests
{
    private sealed class StubBatch : IPresentationBatch
    {
        public string PresentationPath => @"C:\deck.pptx";
        public bool HasTimedOutOperation => false;
        public int? PowerPointProcessId => null;
        public PowerPointProcessIdentity? PowerPointProcessIdentity => null;
        public TimeSpan OperationTimeout => TimeSpan.FromMinutes(1);
        public void Execute(Action<PresentationContext, CancellationToken> operation, CancellationToken cancellationToken = default) => throw new InvalidOperationException("No PowerPoint in unit tests.");
        public T Execute<T>(Func<PresentationContext, CancellationToken, T> operation, CancellationToken cancellationToken = default) => throw new InvalidOperationException("No PowerPoint in unit tests.");
        public void Save(CancellationToken cancellationToken = default) => throw new InvalidOperationException();
        public T TransformPresentationCopy<T>(Func<string, CancellationToken, T> transform, CancellationToken cancellationToken = default) => throw new InvalidOperationException();
        public void UpdatePresentationPath(string presentationPath) { }
        public bool IsPowerPointProcessAlive() => true;
        public void Dispose() { }
    }

    private sealed class FakeDispatcher(Func<string, string, string?, BatchStepOutcome> handler) : IBatchDispatcher
    {
        public List<(string Command, string? Args)> Calls { get; } = [];

        public Action<string>? OnCall { get; set; }

        public BatchStepOutcome Dispatch(string category, string action, string? argumentsJson, string jobId)
        {
            Calls.Add(($"{category}.{action}", argumentsJson));
            OnCall?.Invoke(jobId);
            return handler(category, action, argumentsJson);
        }
    }

    private static readonly BatchStepOutcome Ok = new(true, null, """{"success":true,"shapeIndex":4,"shapeCount":4,"slides":[{"slideId":300}]}""");

    private const string TwoOps = """
        [
          {"id":"box","tool":"shape","action":"add-text-box","args":{"slide_index":1,"left":10,"top":10,"width":200,"height":40,"text":"Привет"}},
          {"id":"size","tool":"textframe","action":"set-font-size","args":{"slide_index":1,"shape_index":"$box.shapeIndex","font_size":24}}
        ]
        """;

    [Fact]
    public void Parse_NormalizesNamesAndRecordsReferences()
    {
        var plan = BatchPlanner.Parse(TwoOps);
        Assert.Empty(plan.Errors);
        Assert.Equal(["shape.add-text-box", "textframe.set-font-size"], plan.Operations.Select(operation => operation.Command));
        Assert.True(plan.Operations[0].Arguments.ContainsKey("slideIndex"));
        Assert.Equal(("box", "shapeIndex"), plan.Operations[1].References["shapeIndex"]);
    }

    [Theory]
    [InlineData("""{"tool":"shape"}""", "expected a JSON array")]
    [InlineData("""[]""", "no operations")]
    [InlineData("""[{"tool":"shape","action":"add-textbox"}]""", "Unknown action 'add-textbox'")]
    [InlineData("""[{"tool":"shape","action":"add-text-box","args":{"slideIdx":1}}]""", "slideIdx")]
    [InlineData("""[{"tool":"batch","action":"run"}]""", "cannot run inside a batch")]
    [InlineData("""[{"tool":"presentation","action":"save"}]""", "cannot run inside a batch")]
    [InlineData("""[{"id":"a","tool":"slide","action":"add-blank"},{"id":"a","tool":"slide","action":"add-blank"}]""", "'a' is used twice")]
    [InlineData("""[{"tool":"slide","action":"delete","args":{"slide_index":"$later.slideIndex"}},{"id":"later","tool":"slide","action":"add-blank"}]""", "not an earlier operation")]
    [InlineData("""[{"tool":"slide","action":"add-blank","extra":1}]""", "unknown field")]
    [InlineData("""[{"command":"slide"}]""", "tool.action")]
    public void Parse_ReportsProblemsBeforeAnythingRuns(string json, string expected)
    {
        var plan = BatchPlanner.Parse(json);
        Assert.Contains(plan.Errors, error => error.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void Parse_AcceptsCommandFormReadToolsAndWrapper()
    {
        var plan = BatchPlanner.Parse("""{"operations":[{"command":"deck_read.summary"},{"tool":"slide","action":"add-blank","args":{"presentation_session_id":"x"}}]}""");
        Assert.Empty(plan.Errors);
        Assert.Equal("deck.summary", plan.Operations[0].Command);
        Assert.Contains(plan.Warnings, warning => warning.Contains("presentation_session_id removed", StringComparison.Ordinal));
    }

    [Fact]
    public void Resolve_SelectsNestedValues()
    {
        var plan = BatchPlanner.Parse("""[{"id":"a","tool":"slide","action":"add-blank"},{"tool":"slide","action":"delete","args":{"slide_index":"$a.slides[0].slideId"}}]""");
        var results = new Dictionary<string, JsonNode?> { ["a"] = JsonNode.Parse(Ok.ResultJson!) };
        var (json, error) = BatchPlanner.Resolve(plan.Operations[1], results);
        Assert.Null(error);
        Assert.Equal(300, JsonNode.Parse(json!)!["slideIndex"]!.GetValue<int>());
        var (_, missing) = BatchPlanner.Resolve(plan.Operations[1], new Dictionary<string, JsonNode?> { ["a"] = new JsonObject() });
        Assert.Contains("no value at 'slides[0].slideId'", missing, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("slide_index", "slideIndex")]
    [InlineData("presentation_session_id", "presentationSessionId")]
    [InlineData("shapeIndex", "shapeIndex")]
    [InlineData("font-size", "fontSize")]
    public void CamelCase(string input, string expected) => Assert.Equal(expected, BatchPlanner.CamelCase(input));

    [Fact]
    public void Run_ExecutesInOrderAndSubstitutesReferences()
    {
        var dispatcher = new FakeDispatcher((_, _, _) => Ok);
        var result = new BatchCommands(dispatcher).Run(new StubBatch(), operations: TwoOps, changeLog: false);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Null(result.ErrorMessage);
        Assert.Equal(("completed", 2, 0), (result.State, result.Completed!.Value, result.Failed!.Value));
        Assert.Contains("\"shapeIndex\":4", dispatcher.Calls[1].Args, StringComparison.Ordinal);
        Assert.Contains("\"text\":\"Привет\"", dispatcher.Calls[0].Args, StringComparison.Ordinal);
        Assert.Equal(4, result.Operations![0].Result!["shapeIndex"]!.GetValue<int>());
        Assert.Null(result.Operations[0].Result!["slides"]);
    }

    [Fact]
    public void Run_StopModeSkipsTheRestAndReportsPartialFailure()
    {
        int call = 0;
        var dispatcher = new FakeDispatcher((_, _, _) => ++call == 2 ? new BatchStepOutcome(false, "Shape index 9 is out of range.", """{"success":false}""") : Ok);
        var json = """[{"tool":"slide","action":"add-blank"},{"tool":"slide","action":"add-blank"},{"tool":"slide","action":"add-blank"}]""";
        var result = new BatchCommands(dispatcher).Run(new StubBatch(), operations: json, changeLog: false);
        Assert.False(result.Success);
        Assert.Equal(["ok", "failed", "skipped"], result.Operations!.Select(step => step.Status));
        Assert.Contains("1 of 3 operations failed; first: op2", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Contains("not undone", result.ErrorMessage, StringComparison.Ordinal);
        Assert.Equal(2, dispatcher.Calls.Count);
    }

    [Fact]
    public void Run_ContinueModeRunsEverything()
    {
        int call = 0;
        var dispatcher = new FakeDispatcher((_, _, _) => ++call == 1 ? new BatchStepOutcome(false, "nope", null) : Ok);
        var json = """[{"tool":"slide","action":"add-blank"},{"tool":"slide","action":"add-blank"}]""";
        var result = new BatchCommands(dispatcher).Run(new StubBatch(), operations: json, mode: "continue", changeLog: false);
        Assert.Equal(["failed", "ok"], result.Operations!.Select(step => step.Status));
        Assert.False(result.Success);
    }

    [Fact]
    public void Run_FailedReferenceTargetFailsTheDependentOperation()
    {
        var dispatcher = new FakeDispatcher((_, _, _) => new BatchStepOutcome(false, "boom", null));
        var result = new BatchCommands(dispatcher).Run(new StubBatch(), operations: TwoOps, mode: "continue", changeLog: false);
        Assert.Equal("failed", result.Operations![1].Status);
        Assert.Contains("has no result", result.Operations[1].Error, StringComparison.Ordinal);
        Assert.Single(dispatcher.Calls);
    }

    [Fact]
    public void Run_CancelledAtOperationBoundary()
    {
        var dispatcher = new FakeDispatcher((_, _, _) => Ok);
        dispatcher.OnCall = jobId => BatchJobs.Find(jobId)!.Cancel();
        var json = """[{"tool":"slide","action":"add-blank"},{"tool":"slide","action":"add-blank"},{"tool":"slide","action":"add-blank"}]""";
        var result = new BatchCommands(dispatcher).Run(new StubBatch(), operations: json, changeLog: false);
        Assert.Equal("cancelled", result.State);
        Assert.Equal(["ok", "cancelled", "cancelled"], result.Operations!.Select(step => step.Status));
        Assert.Single(dispatcher.Calls);
    }

    [Fact]
    public void Start_RefusesASecondJobInTheSameSessionUntilTheFirstFinishes()
    {
        using var gate = new ManualResetEventSlim(false);
        var dispatcher = new FakeDispatcher((_, _, _) => { gate.Wait(TimeSpan.FromSeconds(10)); return Ok; });
        var session = new StubBatch();
        var commands = new BatchCommands(dispatcher);
        var first = commands.Start(session, operations: """[{"tool":"slide","action":"add-blank"}]""", changeLog: false);
        Assert.True(first.Success, first.ErrorMessage);
        var second = commands.Run(session, operations: """[{"tool":"slide","action":"add-blank"}]""", changeLog: false);
        Assert.False(second.Success);
        Assert.Contains("still running", second.ErrorMessage, StringComparison.Ordinal);
        Assert.Same(BatchJobs.Find(first.JobId!), BatchJobs.RunningIn(session));

        gate.Set();
        var deadline = DateTime.UtcNow.AddSeconds(10);
        while (commands.Status(session, first.JobId!).State == "running" && DateTime.UtcNow < deadline)
            Thread.Sleep(20);
        var status = commands.Status(session, first.JobId!);
        Assert.Equal("completed", status.State);
        Assert.True(status.Success);
        Assert.Null(BatchJobs.RunningIn(session));
    }

    [Fact]
    public void Validate_RunsNothing()
    {
        var dispatcher = new FakeDispatcher((_, _, _) => Ok);
        var result = new BatchCommands(dispatcher).Validate(new StubBatch(), operations: TwoOps);
        Assert.True(result.Success);
        Assert.Equal(2, result.Total);
        Assert.Empty(dispatcher.Calls);
    }

    [Fact]
    public void InvalidBatch_ListsEveryProblem()
    {
        var result = new BatchCommands(new FakeDispatcher((_, _, _) => Ok)).Run(new StubBatch(),
            operations: """[{"tool":"shape","action":"nope"},{"tool":"slide","action":"also-nope"}]""");
        Assert.False(result.Success);
        Assert.Equal(2, result.Errors!.Count);
        Assert.Contains("nothing ran", result.ErrorMessage, StringComparison.Ordinal);
    }
}
