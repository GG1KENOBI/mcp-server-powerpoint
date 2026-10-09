// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Net;
using Sbroenne.PowerPointMcp.Service;
using Sbroenne.PowerPointMcp.Service.LocalUi;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.LocalUi;

/// <summary>The optional local UI: loopback only, token, same-origin, read-only commands, embedded assets (no PowerPoint).</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Service")]
[Trait("Feature", "LocalUi")]
[Collection("LocalUi")]
public sealed class LocalUiServerTests : IDisposable
{
    private readonly PowerPointMcpService _service = new();
    private readonly LocalUiServer _server;
    private readonly HttpClient _http = new();
    private readonly string _base;
    private readonly string _token;

    public LocalUiServerTests()
    {
        Environment.SetEnvironmentVariable("PPTMCP_UI", "1");
        try
        {
            _server = LocalUiServer.StartFromEnvironment(_service, _ => { }) ?? throw new InvalidOperationException("The UI did not start.");
        }
        finally
        {
            Environment.SetEnvironmentVariable("PPTMCP_UI", null);
        }
        var url = new Uri(_server.Url);
        _base = url.GetLeftPart(UriPartial.Authority);
        _token = url.Fragment["#token=".Length..];
    }

    public void Dispose()
    {
        _http.Dispose();
        _server.Dispose();
        _service.Dispose();
    }

    private Task<HttpResponseMessage> Get(string path, bool token = true, string? origin = null)
    {
        var request = new HttpRequestMessage(HttpMethod.Get, _base + path);
        if (token)
            request.Headers.Add("X-PptMcp-Token", _token);
        if (origin is not null)
            request.Headers.Add("Origin", origin);
        return _http.SendAsync(request);
    }

    [Fact]
    public void IsOffUnlessEnabled()
    {
        Assert.Null(LocalUiServer.StartFromEnvironment(_service, _ => { }));
        Assert.StartsWith("http://localhost:", _server.Url, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ServesEmbeddedPageWithStrictHeadersAndNoCors()
    {
        using var page = await Get("/", token: false);
        var html = await page.Content.ReadAsStringAsync();

        Assert.Equal(HttpStatusCode.OK, page.StatusCode);
        Assert.Contains("PowerPoint MCP Review", html, StringComparison.Ordinal);
        Assert.DoesNotContain("http", html.Replace("http-equiv", "", StringComparison.Ordinal), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("default-src 'self'", page.Headers.GetValues("Content-Security-Policy").Single(), StringComparison.Ordinal);
        Assert.False(page.Headers.Contains("Access-Control-Allow-Origin"));
        using var script = await Get("/app.js", token: false);
        Assert.Equal(HttpStatusCode.OK, script.StatusCode);
    }

    [Fact]
    public async Task ApiNeedsTheTokenAndRejectsOtherOrigins()
    {
        using var missing = await Get("/api/sessions", token: false);
        using var crossOrigin = await Get("/api/sessions", origin: "http://evil.example");
        using var ok = await Get("/api/sessions", origin: _base);

        Assert.Equal(HttpStatusCode.Unauthorized, missing.StatusCode);
        Assert.Equal(HttpStatusCode.Forbidden, crossOrigin.StatusCode);
        Assert.Equal(HttpStatusCode.OK, ok.StatusCode);
        Assert.Contains("\"sessions\":[]", await ok.Content.ReadAsStringAsync(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task OnlyReadOnlyCommandsAndPreviewFilesAreAllowed()
    {
        using var write = await Get("/api/command?command=deck.replace-text&session=x");
        using var file = await Get("/api/file?path=" + Uri.EscapeDataString(Path.Combine(Path.GetTempPath(), "secret.png")));
        using var post = await _http.SendAsync(new HttpRequestMessage(HttpMethod.Post, _base + "/api/sessions") { Headers = { { "X-PptMcp-Token", _token } } });

        Assert.Equal(HttpStatusCode.Forbidden, write.StatusCode);
        Assert.Equal(HttpStatusCode.NotFound, file.StatusCode);
        Assert.Equal(HttpStatusCode.MethodNotAllowed, post.StatusCode);
    }

    [Fact]
    public async Task CapabilitiesIncludeTheUiAddress()
    {
        using var response = await Get("/api/capabilities");
        var body = await response.Content.ReadAsStringAsync();

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Contains("\"workflows\"", body, StringComparison.Ordinal);
    }
}
