// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.IO.Pipelines;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Client;
using Sbroenne.PowerPointMcp.Generated;
using Xunit.Abstractions;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Integration;

/// <summary>
/// Near-miss tool calls: strict by default (rejected with a precise suggestion), normalized to the
/// canonical names when <c>PPTMCP_LENIENT_ARGUMENTS</c> is enabled. No PowerPoint required: calls
/// use an unknown session, so a call that passes validation ends at the service's own
/// "not found" result — identical to the canonical call's result.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "McpProtocol")]
public sealed class McpToolCallNormalizationTests : IAsyncLifetime, IAsyncDisposable
{
    private const string MissingSession = "missing-session";

    private readonly ITestOutputHelper _output;
    private readonly Pipe _clientToServerPipe = new();
    private readonly Pipe _serverToClientPipe = new();
    private readonly CancellationTokenSource _cts = new();
    private McpClient? _client;
    private Task? _serverTask;

    public McpToolCallNormalizationTests(ITestOutputHelper output)
    {
        _output = output;
    }

    public async Task InitializeAsync()
    {
        ToolCallNormalizer.LenientNames = false;
        (_client, _serverTask) = await ProgramTransportTestHost.StartAsync(
            _clientToServerPipe,
            _serverToClientPipe,
            "NormalizationTestClient",
            _cts.Token);
    }

    public async Task DisposeAsync() => await DisposeAsyncCore();

    async ValueTask IAsyncDisposable.DisposeAsync()
    {
        await DisposeAsyncCore();
        GC.SuppressFinalize(this);
    }

    private async Task DisposeAsyncCore()
    {
        ToolCallNormalizer.LenientNames = false;
        await ProgramTransportTestHost.StopAsync(
            _client,
            _clientToServerPipe,
            _serverToClientPipe,
            _serverTask,
            _output);

        _cts.Dispose();
    }

    private static Dictionary<string, object?> TextBox(string action, string sessionKey = "presentation_session_id") => new()
    {
        ["action"] = action,
        [sessionKey] = MissingSession,
        ["slide_index"] = 1,
        ["left"] = 10,
        ["top"] = 20,
        ["width"] = 300,
        ["height"] = 40,
        ["text"] = "Hello",
    };

    [Theory]
    [InlineData("1", true)]
    [InlineData("true", true)]
    [InlineData(" TRUE ", true)]
    [InlineData("0", false)]
    [InlineData("yes", false)]
    [InlineData(null, false)]
    public void LenientMode_IsOptInThroughTheEnvironmentVariable(string? value, bool enabled)
    {
        Assert.Equal(enabled, ToolCallNormalizer.IsEnabled(value));
    }

    [Fact]
    public async Task CanonicalShapeCall_PassesValidationUnchanged()
    {
        Assert.Equal($"Session '{MissingSession}' not found", await ErrorMessageAsync("shape", TextBox("add-text-box")));
    }

    [Fact]
    public async Task StrictMode_RejectsNearMissActionWithTheCanonicalName()
    {
        var message = await ErrorMessageAsync("shape", TextBox("add_text_box"));

        Assert.Contains("Unknown action \"add_text_box\"", message, StringComparison.Ordinal);
        Assert.Contains("Did you mean \"add-text-box\"?", message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task StrictMode_RejectsLegacySessionNameWithTheCanonicalName()
    {
        var message = await ErrorMessageAsync("presentation", new() { ["action"] = "close", ["session_id"] = MissingSession });

        Assert.Contains("Unknown parameter 'session_id'", message, StringComparison.Ordinal);
        Assert.Contains("Did you mean 'presentation_session_id'?", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("add_text_box")]
    [InlineData("addTextBox")]
    [InlineData("AddTextBox")]
    [InlineData("ADD-TEXT-BOX")]
    [InlineData("add text box")]
    public async Task LenientMode_ActionSpellingsNormalizeToTheCanonicalAction(string action)
    {
        var canonical = await RawResultAsync("shape", TextBox("add-text-box"));
        ToolCallNormalizer.LenientNames = true;

        Assert.Equal(canonical, await RawResultAsync("shape", TextBox(action)));
    }

    [Theory]
    [InlineData("sessionId")]
    [InlineData("session_id")]
    [InlineData("presentationSessionId")]
    public async Task LenientMode_SessionAndCamelCaseParametersNormalize(string sessionKey)
    {
        var canonical = await RawResultAsync("shape", TextBox("add-text-box"));
        ToolCallNormalizer.LenientNames = true;
        var arguments = TextBox("addTextBox", sessionKey);
        arguments.Remove("slide_index");
        arguments["slideIndex"] = 1;

        Assert.Equal(canonical, await RawResultAsync("shape", arguments));
    }

    [Fact]
    public async Task LenientMode_PresentationParametersNormalize()
    {
        var canonical = await RawResultAsync("presentation", new()
        {
            ["action"] = "save-as",
            ["presentation_session_id"] = MissingSession,
            ["targetPath"] = @"C:\temp\out.pptx",
        });
        ToolCallNormalizer.LenientNames = true;
        var lenient = await RawResultAsync("presentation", new()
        {
            ["action"] = "save_as",
            ["session_id"] = MissingSession,
            ["target_path"] = @"C:\temp\out.pptx",
        });

        Assert.Contains($"Unknown presentation_session_id: {MissingSession}", canonical, StringComparison.Ordinal);
        Assert.Equal(canonical, lenient);
    }

    [Fact]
    public async Task LenientMode_CanonicalAndAliasForTheSameParameter_IsRejectedNotGuessed()
    {
        ToolCallNormalizer.LenientNames = true;
        var arguments = TextBox("add-text-box");
        arguments["sessionId"] = "other-session";

        var message = await ErrorMessageAsync("shape", arguments);

        Assert.Contains("Parameter 'sessionId' duplicates 'presentation_session_id'", message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UnknownAction_SuggestsClosestActionAndItsRequiredParameters()
    {
        var message = await ErrorMessageAsync("shape", new()
        {
            ["action"] = "add_text_bx",
            ["presentation_session_id"] = MissingSession,
        });

        Assert.Contains("Unknown action \"add_text_bx\"", message, StringComparison.Ordinal);
        Assert.Contains("Did you mean \"add-text-box\"", message, StringComparison.Ordinal);
        Assert.Contains(
            "Required parameters for \"add-text-box\": presentation_session_id, slide_index, left, top, width, height, text.",
            message,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingDimensions_ListsEveryMissingParameterWithAnExample()
    {
        var message = await ErrorMessageAsync("shape", new()
        {
            ["action"] = "add-text-box",
            ["presentation_session_id"] = MissingSession,
            ["slide_index"] = 1,
            ["text"] = "Hello",
        });

        Assert.Contains(
            "left, top, width, height are required for shape action \"add-text-box\".",
            message,
            StringComparison.Ordinal);
        var example = JsonNode.Parse(message[(message.IndexOf("Example: ", StringComparison.Ordinal) + "Example: ".Length)..])!.AsObject();
        Assert.Equal("add-text-box", example["action"]!.GetValue<string>());
        foreach (var name in new[] { "presentation_session_id", "slide_index", "left", "top", "width", "height", "text" })
        {
            Assert.True(example.ContainsKey(name), $"Example is missing {name}: {example}");
        }
    }

    [Fact]
    public async Task UnknownParameter_SuggestsClosestValidParameter()
    {
        var message = await ErrorMessageAsync("shape", new()
        {
            ["action"] = "add-text-box",
            ["presentation_session_id"] = MissingSession,
            ["slide_idx"] = 1,
        });

        Assert.Contains("Unknown parameter 'slide_idx'", message, StringComparison.Ordinal);
        Assert.Contains("Did you mean 'slide_index'?", message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task InapplicableParameter_ErrorListsTheActionsValidParameters()
    {
        var message = await ErrorMessageAsync("presentation", new()
        {
            ["action"] = "get-final",
            ["presentation_session_id"] = MissingSession,
            ["isFinal"] = false,
        });

        Assert.Contains("not valid for action 'get-final'", message, StringComparison.Ordinal);
        Assert.Contains("Valid parameters for \"get-final\": presentation_session_id.", message, StringComparison.Ordinal);
    }

    /// <summary>Normalization is only safe if no two canonical names share a folded key.</summary>
    [Fact]
    public async Task AllTools_CanonicalNamesAreUnambiguousAfterFolding()
    {
        foreach (var tool in await _client!.ListToolsAsync(cancellationToken: _cts.Token))
        {
            var metadata = ToolCallNormalizer.GetMetadata(tool.ProtocolTool);
            Assert.NotEmpty(metadata.Actions);
            AssertUniqueKeys(tool.Name, metadata.Actions);
            AssertUniqueKeys(tool.Name, metadata.Parameters);
            foreach (var action in metadata.Actions)
            {
                Assert.Equal(action, metadata.ResolveAction(action));
            }
        }

        static void AssertUniqueKeys(string tool, IReadOnlyList<string> names)
        {
            var collisions = names.GroupBy(ToolCallNormalizer.Fold).Where(group => group.Count() > 1).ToArray();
            Assert.True(collisions.Length == 0, $"{tool}: ambiguous names {string.Join("; ", collisions.Select(g => string.Join("/", g)))}");
        }
    }

    /// <summary>
    /// Required/valid parameters derived from each generated tool's schema must agree with the
    /// generated service validators, so the pre-dispatch check never rejects a call the backend
    /// would accept (or vice versa).
    /// </summary>
    [Fact]
    public async Task GeneratedTools_DerivedRequirementsMatchServiceValidators()
    {
        var registryTypes = typeof(ServiceRegistry).GetNestedTypes(BindingFlags.Public);
        var checkedActions = 0;
        foreach (var tool in await _client!.ListToolsAsync(cancellationToken: _cts.Token))
        {
            if (tool.Name == "presentation" || tool.Name.EndsWith("_read", StringComparison.Ordinal))
                continue;

            var metadata = ToolCallNormalizer.GetMetadata(tool.ProtocolTool);
            var properties = tool.JsonSchema.GetProperty("properties");
            var registry = Assert.Single(registryTypes, type => ToolCallNormalizer.Fold(type.Name) == tool.Name);
            var validateArguments = registry.GetMethod("ValidateActionArguments", BindingFlags.Public | BindingFlags.Static)!;

            foreach (var action in metadata.Actions)
            {
                var valid = metadata.ValidParameters(action).Where(name => name != "presentation_session_id").ToArray();
                ServiceRegistry.ValidateMcpActionParameters(tool.Name, action, valid);

                var required = metadata.RequiredParameters(action).Where(name => name != "presentation_session_id").ToArray();
                Assert.All(required, name => Assert.Contains(name, valid));

                var error = Invoke(validateArguments, action, BuildServiceArguments(properties, required));
                Assert.True(
                    error == null || !error.Contains("required", StringComparison.OrdinalIgnoreCase),
                    $"{tool.Name} {action}: service still requires something the schema does not mark required: {error}");

                foreach (var omitted in required)
                {
                    var omittedError = Invoke(validateArguments, action, BuildServiceArguments(properties, required.Where(name => name != omitted)));
                    Assert.True(
                        omittedError != null && omittedError.Contains(ToCamelCase(omitted), StringComparison.Ordinal),
                        $"{tool.Name} {action}: schema marks {omitted} required but the service accepted it missing ({omittedError ?? "no error"}).");
                }
                checkedActions++;
            }
        }
        _output.WriteLine($"Checked {checkedActions} generated actions.");
        Assert.True(checkedActions > 100);
    }

    private static string? Invoke(MethodInfo validateArguments, string action, string argsJson)
    {
        try
        {
            validateArguments.Invoke(null, [action, argsJson]);
            return null;
        }
        catch (TargetInvocationException ex)
        {
            return ex.InnerException?.Message ?? ex.Message;
        }
    }

    private static string BuildServiceArguments(JsonElement properties, IEnumerable<string> names)
    {
        var arguments = new JsonObject();
        foreach (var name in names)
        {
            arguments[ToCamelCase(name)] = DummyValue(properties.GetProperty(name));
        }
        return arguments.ToJsonString();
    }

    private static JsonNode DummyValue(JsonElement schema) => PrimaryType(schema) switch
    {
        "integer" => 1,
        "number" => 1,
        "boolean" => true,
        "array" => PrimaryType(schema.GetProperty("items")) == "string" ? new JsonArray("a", "b") : new JsonArray(1, 2),
        _ => "x",
    };

    private static string PrimaryType(JsonElement schema)
    {
        if (!schema.TryGetProperty("type", out var type))
            return "string";
        return type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Select(item => item.GetString()).First(item => item != "null")!
            : type.GetString()!;
    }

    private static string ToCamelCase(string snake)
    {
        var parts = snake.Split('_');
        return parts[0] + string.Concat(parts.Skip(1).Select(part => char.ToUpperInvariant(part[0]) + part[1..]));
    }

    private async Task<string> RawResultAsync(string tool, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(tool, arguments, cancellationToken: _cts.Token);
        return Assert.IsType<JsonElement>(result.StructuredContent).GetRawText();
    }

    private async Task<string> ErrorMessageAsync(string tool, Dictionary<string, object?> arguments)
    {
        var result = await _client!.CallToolAsync(tool, arguments, cancellationToken: _cts.Token);
        Assert.True(result.IsError);
        var message = Assert.IsType<JsonElement>(result.StructuredContent).GetProperty("errorMessage").GetString()!;
        _output.WriteLine(message);
        return message;
    }
}
