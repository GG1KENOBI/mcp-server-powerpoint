using System.Runtime.CompilerServices;

namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>
/// Rendered-preview cache. An entry is reused only when the slide fingerprint, the session's
/// change counter, and the requested width all match and the file is under ten minutes old. The
/// service bumps the counter after every mutating command, so edits through this server always
/// invalidate. Edits made directly in PowerPoint that change neither geometry nor text are not
/// visible to the fingerprint; pass refresh=true after such edits.
/// </summary>
public static class PreviewCache
{
    /// <summary>How long a cached preview may be reused.</summary>
    public static readonly TimeSpan MaxAge = TimeSpan.FromMinutes(10);

    private static readonly ConditionalWeakTable<object, SessionState> States = new();

    private sealed class SessionState
    {
        public long Generation;

        public string Token { get; } = Guid.NewGuid().ToString("N")[..12];
    }

    /// <summary>Folder for previews: PPTMCP_PREVIEW_DIR or %LOCALAPPDATA%\PowerPointMcp\previews.</summary>
    public static string Directory =>
        Environment.GetEnvironmentVariable("PPTMCP_PREVIEW_DIR") is { Length: > 0 } custom
            ? custom
            : Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerPointMcp", "previews");

    /// <summary>Records that the session's presentation changed.</summary>
    public static void MarkChanged(object session)
    {
        ArgumentNullException.ThrowIfNull(session);
        Interlocked.Increment(ref States.GetOrCreateValue(session).Generation);
    }

    /// <summary>The session's change counter.</summary>
    public static long Generation(object session)
    {
        ArgumentNullException.ThrowIfNull(session);
        return Interlocked.Read(ref States.GetOrCreateValue(session).Generation);
    }

    /// <summary>The file a preview with these inputs is cached at.</summary>
    public static string PathFor(object session, int slideId, int width, string fingerprint)
    {
        var state = States.GetOrCreateValue(session);
        return Path.Combine(Directory, $"{state.Token}-s{slideId}-w{width}-{fingerprint}-g{Interlocked.Read(ref state.Generation)}.png");
    }

    /// <summary>Whether a usable cached file exists.</summary>
    public static bool IsFresh(string path) =>
        File.Exists(path) && DateTime.UtcNow - File.GetLastWriteTimeUtc(path) < MaxAge && new FileInfo(path).Length > 0;

    /// <summary>Deletes cached previews older than <see cref="MaxAge"/> (or all when <paramref name="all"/>); returns the count.</summary>
    public static int Prune(bool all = false)
    {
        if (!System.IO.Directory.Exists(Directory))
            return 0;
        int removed = 0;
        foreach (var file in System.IO.Directory.EnumerateFiles(Directory, "*.png"))
        {
            if (!all && DateTime.UtcNow - File.GetLastWriteTimeUtc(file) < MaxAge)
                continue;
            try
            {
                File.Delete(file);
                removed++;
            }
            catch (IOException)
            {
                // In use by a reader; it will be pruned next time.
            }
            catch (UnauthorizedAccessException)
            {
                // Not ours to delete.
            }
        }
        return removed;
    }
}
