using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Sbroenne.PowerPointMcp.Core.Catalog;
using Sbroenne.PowerPointMcp.Core.Preview;

namespace Sbroenne.PowerPointMcp.Service.LocalUi;

/// <summary>
/// Optional local review UI (PPTMCP_UI=1): a small web page served by the process that already
/// hosts <see cref="PowerPointMcpService"/>, so it shows the same sessions the agent uses. It is
/// not a second backend: every API call goes through the service's own dispatch, and only
/// read-only commands are allowed.
/// </summary>
/// <remarks>
/// Security: listens on localhost only; rejects non-loopback clients, Host headers other than
/// localhost/127.0.0.1 on its port (DNS rebinding), and any Origin other than its own; never
/// sends CORS headers; API calls need a random per-run token (sent in a header, delivered in the
/// URL fragment so it never reaches logs or Referer); serves files only from the preview folder;
/// all page assets are embedded in the assembly (no CDN). Content-Security-Policy allows only
/// same-origin resources.
/// </remarks>
public sealed class LocalUiServer : IDisposable
{
    private const string TokenHeader = "X-PptMcp-Token";
    private static readonly JsonSerializerOptions Json = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull };

    private readonly PowerPointMcpService _service;
    private readonly HttpListener _listener = new();
    private readonly CancellationTokenSource _stop = new();
    private readonly byte[] _token;
    private readonly int _port;
    private readonly Task _loop;

    private LocalUiServer(PowerPointMcpService service, int port)
    {
        _service = service;
        _port = port;
        _token = RandomNumberGenerator.GetBytes(24);
        _listener.Prefixes.Add($"http://localhost:{port}/");
        _listener.Start();
        _loop = Task.Run(LoopAsync);
    }

    /// <summary>The running UI, if any.</summary>
    public static LocalUiServer? Current { get; private set; }

    /// <summary>The URL to open (the token is in the fragment).</summary>
    public string Url => $"http://localhost:{_port}/#token={Convert.ToHexString(_token).ToLowerInvariant()}";

    /// <summary>Starts the UI when PPTMCP_UI is 1/true; returns null (and starts nothing) otherwise.</summary>
    public static LocalUiServer? StartFromEnvironment(PowerPointMcpService service, Action<string> log)
    {
        ArgumentNullException.ThrowIfNull(service);
        ArgumentNullException.ThrowIfNull(log);
        var enabled = Environment.GetEnvironmentVariable("PPTMCP_UI");
        if (enabled is null || !(enabled == "1" || enabled.Equals("true", StringComparison.OrdinalIgnoreCase)))
            return null;
        int port = int.TryParse(Environment.GetEnvironmentVariable("PPTMCP_UI_PORT"), out var configured) && configured is > 1023 and < 65536 ? configured : FreePort();
        try
        {
            var server = new LocalUiServer(service, port);
            Current = server;
            log($"PowerPoint MCP local UI: {server.Url}");
            return server;
        }
        catch (HttpListenerException ex)
        {
            log($"The local UI could not start on port {port}: {ex.Message}");
            return null;
        }
    }

    private static int FreePort()
    {
        var probe = new TcpListener(IPAddress.Loopback, 0);
        probe.Start();
        try
        {
            return ((IPEndPoint)probe.LocalEndpoint).Port;
        }
        finally
        {
            probe.Stop();
        }
    }

    private async Task LoopAsync()
    {
        while (!_stop.IsCancellationRequested)
        {
            HttpListenerContext context;
            try
            {
                context = await _listener.GetContextAsync().ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException or InvalidOperationException)
            {
                return;
            }
            _ = Task.Run(() => HandleAsync(context));
        }
    }

    private async Task HandleAsync(HttpListenerContext context)
    {
        var request = context.Request;
        var response = context.Response;
        try
        {
            AddSecurityHeaders(response);
            if (!IsAllowed(request, out var reason))
            {
                await WriteAsync(response, 403, "text/plain; charset=utf-8", Encoding.UTF8.GetBytes(reason)).ConfigureAwait(false);
                return;
            }
            var path = request.Url?.AbsolutePath ?? "/";
            if (request.HttpMethod != "GET")
            {
                await WriteAsync(response, 405, "text/plain; charset=utf-8", "Only GET is supported."u8.ToArray()).ConfigureAwait(false);
                return;
            }
            if (!path.StartsWith("/api/", StringComparison.Ordinal))
            {
                await ServeAssetAsync(response, path).ConfigureAwait(false);
                return;
            }
            if (!HasToken(request))
            {
                await WriteJsonAsync(response, 401, new { success = false, errorMessage = "Missing or wrong token; open the URL printed by the server." }).ConfigureAwait(false);
                return;
            }
            await HandleApiAsync(request, response, path).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is HttpListenerException or IOException or ObjectDisposedException)
        {
            // The browser went away; nothing to report.
        }
        finally
        {
            try
            {
                response.Close();
            }
            catch (Exception ex) when (ex is HttpListenerException or ObjectDisposedException or InvalidOperationException)
            {
                // Already closed.
            }
        }
    }

    private bool IsAllowed(HttpListenerRequest request, out string reason)
    {
        reason = "";
        if (!IPAddress.IsLoopback(request.RemoteEndPoint.Address))
        {
            reason = "Only this computer may connect.";
            return false;
        }
        var host = request.Headers["Host"] ?? "";
        if (host != $"localhost:{_port}" && host != $"127.0.0.1:{_port}")
        {
            reason = "Unexpected Host header.";
            return false;
        }
        var origin = request.Headers["Origin"];
        if (origin is not null && origin != $"http://localhost:{_port}" && origin != $"http://127.0.0.1:{_port}")
        {
            reason = "Cross-origin requests are not allowed.";
            return false;
        }
        return true;
    }

    private bool HasToken(HttpListenerRequest request)
    {
        var supplied = request.Headers[TokenHeader];
        if (supplied is null)
            return false;
        byte[] bytes;
        try
        {
            bytes = Convert.FromHexString(supplied);
        }
        catch (FormatException)
        {
            return false;
        }
        return CryptographicOperations.FixedTimeEquals(bytes, _token);
    }

    private static void AddSecurityHeaders(HttpListenerResponse response)
    {
        response.Headers["Content-Security-Policy"] = "default-src 'self'; img-src 'self' blob:; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'";
        response.Headers["X-Content-Type-Options"] = "nosniff";
        response.Headers["Referrer-Policy"] = "no-referrer";
        response.Headers["Cache-Control"] = "no-store";
        response.Headers["X-Frame-Options"] = "DENY";
    }

    private async Task HandleApiAsync(HttpListenerRequest request, HttpListenerResponse response, string path)
    {
        var query = request.QueryString;
        switch (path)
        {
            case "/api/sessions":
                await WriteJsonAsync(response, 200, new { success = true, sessions = _service.Sessions.List() }).ConfigureAwait(false);
                return;
            case "/api/capabilities":
                var report = Capabilities.Build("overview", typeof(LocalUiServer).Assembly.GetName().Version?.ToString() ?? "0.0.0", _service.SessionCount, Url);
                await WriteJsonAsync(response, 200, new { success = true, report }).ConfigureAwait(false);
                return;
            case "/api/command":
                await RunCommandAsync(response, query["command"], query["session"], query["args"]).ConfigureAwait(false);
                return;
            case "/api/file":
                await ServePreviewFileAsync(response, query["path"]).ConfigureAwait(false);
                return;
            default:
                await WriteJsonAsync(response, 404, new { success = false, errorMessage = "Unknown API path." }).ConfigureAwait(false);
                return;
        }
    }

    private async Task RunCommandAsync(HttpListenerResponse response, string? command, string? session, string? args)
    {
        var parts = (command ?? "").Split('.', 2);
        if (parts.Length != 2 || !CommandCatalog.IsReadOnly(parts[0], parts[1]))
        {
            await WriteJsonAsync(response, 403, new { success = false, errorMessage = $"'{command}' is not a read-only command; the local UI only reads." }).ConfigureAwait(false);
            return;
        }
        var result = await _service.ProcessAsync(new ServiceRequest { Command = command!, SessionId = session, Args = string.IsNullOrWhiteSpace(args) ? null : args, Source = "ui" }).ConfigureAwait(false);
        var body = result.Success && result.Result is not null
            ? Encoding.UTF8.GetBytes(result.Result)
            : JsonSerializer.SerializeToUtf8Bytes(new { success = false, errorMessage = result.ErrorMessage }, Json);
        await WriteAsync(response, 200, "application/json; charset=utf-8", body).ConfigureAwait(false);
    }

    private static async Task ServePreviewFileAsync(HttpListenerResponse response, string? path)
    {
        string full;
        try
        {
            full = Path.GetFullPath(path ?? "");
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            await WriteJsonAsync(response, 400, new { success = false, errorMessage = "Bad path." }).ConfigureAwait(false);
            return;
        }
        var root = Path.GetFullPath(PreviewCache.Directory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !full.EndsWith(".png", StringComparison.OrdinalIgnoreCase) || !File.Exists(full))
        {
            await WriteJsonAsync(response, 404, new { success = false, errorMessage = "Only preview images can be shown." }).ConfigureAwait(false);
            return;
        }
        await WriteAsync(response, 200, "image/png", await File.ReadAllBytesAsync(full).ConfigureAwait(false)).ConfigureAwait(false);
    }

    private static async Task ServeAssetAsync(HttpListenerResponse response, string path)
    {
        var (name, type) = path switch
        {
            "/" or "/index.html" => ("index.html", "text/html; charset=utf-8"),
            "/app.js" => ("app.js", "text/javascript; charset=utf-8"),
            "/app.css" => ("app.css", "text/css; charset=utf-8"),
            _ => (null, null),
        };
        using var stream = name is null ? null : Assembly.GetExecutingAssembly().GetManifestResourceStream($"LocalUi/{name}");
        if (stream is null)
        {
            await WriteAsync(response, 404, "text/plain; charset=utf-8", "Not found."u8.ToArray()).ConfigureAwait(false);
            return;
        }
        using var buffer = new MemoryStream();
        await stream.CopyToAsync(buffer).ConfigureAwait(false);
        await WriteAsync(response, 200, type!, buffer.ToArray()).ConfigureAwait(false);
    }

    private static Task WriteJsonAsync(HttpListenerResponse response, int status, object payload) =>
        WriteAsync(response, status, "application/json; charset=utf-8", JsonSerializer.SerializeToUtf8Bytes(payload, Json));

    private static async Task WriteAsync(HttpListenerResponse response, int status, string contentType, byte[] body)
    {
        response.StatusCode = status;
        response.ContentType = contentType;
        response.ContentLength64 = body.Length;
        await response.OutputStream.WriteAsync(body).ConfigureAwait(false);
    }

    /// <summary>Stops listening.</summary>
    public void Dispose()
    {
        if (Current == this)
            Current = null;
        _stop.Cancel();
        try
        {
            _listener.Stop();
            _listener.Close();
        }
        catch (ObjectDisposedException)
        {
            // Already stopped.
        }
        try
        {
            _loop.Wait(TimeSpan.FromSeconds(2));
        }
        catch (AggregateException)
        {
            // The loop ends when the listener stops.
        }
        _stop.Dispose();
    }
}
