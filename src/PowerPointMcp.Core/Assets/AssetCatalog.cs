using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>One image in the local asset catalog.</summary>
public sealed class AssetEntry
{
    /// <summary>Full path of the file.</summary>
    public required string Path { get; set; }

    /// <summary>png, jpeg, gif, bmp, tiff, webp, svg, emf, or wmf.</summary>
    public required string Format { get; set; }

    /// <summary>Displayed pixel width (EXIF rotation applied); 0 for vector images without a size.</summary>
    public int Width { get; set; }

    /// <summary>Displayed pixel height.</summary>
    public int Height { get; set; }

    /// <summary>landscape, portrait, or square.</summary>
    public string Orientation { get; set; } = "landscape";

    /// <summary>File size in bytes.</summary>
    public long Bytes { get; set; }

    /// <summary>First 16 hex digits of the SHA-256 of the file (equal hashes are duplicates).</summary>
    public required string Hash { get; set; }

    /// <summary>Last write time (UTC, ISO 8601) when scanned.</summary>
    public required string Modified { get; set; }

    /// <summary>Search tags (lower case).</summary>
    public List<string> Tags { get; set; } = [];

    /// <summary>Free-text description used by search and as default alt text.</summary>
    public string? Description { get; set; }

    /// <summary>Attribution or license text placed with the picture.</summary>
    public string? Attribution { get; set; }
}

/// <summary>A search hit with its score.</summary>
public sealed record AssetHit(AssetEntry Asset, int Score);

/// <summary>What a folder scan found.</summary>
public sealed record AssetScanSummary(int Added, int Updated, int Unchanged, int Removed, int Skipped, IReadOnlyList<string> Notes);

/// <summary>
/// A local catalog of image files (JSON in PPTMCP_ASSETS_DIR or %LOCALAPPDATA%\PowerPointMcp\assets).
/// Scanning reads headers and hashes only; nothing is uploaded, downloaded, or generated.
/// Search ranks file name, tags, description, and folder words; orientation and minimum size
/// filter. Not thread-safe: the service runs one command per session at a time and the catalog
/// file is written atomically.
/// </summary>
public sealed class AssetCatalog
{
    private const int MaxFilesPerScan = 20_000;

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private readonly Dictionary<string, AssetEntry> _entries;

    private AssetCatalog(string file, IEnumerable<AssetEntry> entries)
    {
        File = file;
        _entries = entries.ToDictionary(entry => entry.Path, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>The catalog file.</summary>
    public string File { get; }

    /// <summary>All entries, by path.</summary>
    public IReadOnlyCollection<AssetEntry> Entries => _entries.Values;

    /// <summary>The default catalog folder.</summary>
    public static string DefaultDirectory =>
        Environment.GetEnvironmentVariable("PPTMCP_ASSETS_DIR") is { Length: > 0 } custom
            ? custom
            : System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PowerPointMcp", "assets");

    /// <summary>Loads the catalog (empty when the file does not exist yet).</summary>
    public static AssetCatalog Load(string? directory = null)
    {
        var file = System.IO.Path.Combine(directory ?? DefaultDirectory, "catalog.json");
        if (!System.IO.File.Exists(file))
            return new AssetCatalog(file, []);
        try
        {
            var entries = JsonSerializer.Deserialize<List<AssetEntry>>(System.IO.File.ReadAllText(file), Options) ?? [];
            return new AssetCatalog(file, entries.Where(entry => !string.IsNullOrWhiteSpace(entry.Path)));
        }
        catch (JsonException ex)
        {
            throw new InvalidDataException($"The asset catalog {file} is not valid JSON ({ex.Message}). Move it away to start a new catalog.");
        }
    }

    /// <summary>Writes the catalog atomically.</summary>
    public void Save()
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(File)!);
        var temp = File + ".tmp";
        System.IO.File.WriteAllText(temp, JsonSerializer.Serialize(_entries.Values.OrderBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase).ToList(), Options));
        System.IO.File.Move(temp, File, overwrite: true);
    }

    /// <summary>
    /// Adds or refreshes every supported image under <paramref name="folder"/> and drops entries
    /// under it whose files are gone. Unchanged files (same size and time) are not re-read.
    /// </summary>
    public AssetScanSummary Scan(string folder, bool recursive)
    {
        var root = System.IO.Path.GetFullPath(folder);
        if (!Directory.Exists(root))
            throw new ArgumentException($"The folder '{root}' does not exist.");
        int added = 0, updated = 0, unchanged = 0, skipped = 0;
        var notes = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var files = Directory.EnumerateFiles(root, "*", new EnumerationOptions
        {
            RecurseSubdirectories = recursive,
            IgnoreInaccessible = true,
            AttributesToSkip = FileAttributes.Hidden | FileAttributes.System,
        }).Where(path => ImageHeaderReader.SupportedExtensions.Contains(System.IO.Path.GetExtension(path)));
        foreach (var path in files)
        {
            if (seen.Count >= MaxFilesPerScan)
            {
                notes.Add($"Stopped after {MaxFilesPerScan} files; scan subfolders separately.");
                break;
            }
            seen.Add(path);
            var info = new FileInfo(path);
            var modified = info.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture);
            if (_entries.TryGetValue(path, out var existing) && existing.Bytes == info.Length && existing.Modified == modified)
            {
                unchanged++;
                continue;
            }
            var entry = Read(path, info, modified);
            if (entry is null)
            {
                skipped++;
                continue;
            }
            if (existing is not null)
            {
                entry.Tags = existing.Tags;
                entry.Description = existing.Description;
                entry.Attribution = existing.Attribution;
                updated++;
            }
            else
            {
                added++;
            }
            _entries[path] = entry;
        }
        var prefix = root.EndsWith(System.IO.Path.DirectorySeparatorChar) ? root : root + System.IO.Path.DirectorySeparatorChar;
        var gone = _entries.Keys.Where(path => path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && (recursive || string.Equals(System.IO.Path.GetDirectoryName(path), root, StringComparison.OrdinalIgnoreCase))
            && !seen.Contains(path)).ToList();
        foreach (var path in gone)
            _entries.Remove(path);
        if (skipped > 0)
            notes.Add($"{skipped} file(s) had a supported extension but no readable image header.");
        return new AssetScanSummary(added, updated, unchanged, gone.Count, skipped, notes);
    }

    /// <summary>Reads one file into an entry, or null when it is not a readable image.</summary>
    public static AssetEntry? Read(string path, FileInfo info, string modified)
    {
        ArgumentNullException.ThrowIfNull(info);
        ImageHeaderInfo? header;
        try
        {
            header = ImageHeaderReader.ReadFile(path);
        }
        catch (IOException)
        {
            return null;
        }
        if (header is null)
            return null;
        return new AssetEntry
        {
            Path = path,
            Format = header.Format,
            Width = header.DisplayWidth,
            Height = header.DisplayHeight,
            Orientation = OrientationOf(header.DisplayWidth, header.DisplayHeight),
            Bytes = info.Length,
            Hash = HashOf(path),
            Modified = modified,
        };
    }

    /// <summary>landscape, portrait, or square (within 5%).</summary>
    public static string OrientationOf(int width, int height) =>
        width <= 0 || height <= 0 ? "landscape"
        : Math.Abs(width - height) <= 0.05 * Math.Max(width, height) ? "square"
        : width > height ? "landscape" : "portrait";

    /// <summary>First 16 hex digits of the file's SHA-256.</summary>
    public static string HashOf(string path)
    {
        using var stream = System.IO.File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream))[..16].ToLowerInvariant();
    }

    /// <summary>
    /// Ranks entries for a query: each query word scores 3 in the file name or tags, 2 in the
    /// description, 1 in the folder path. All words must match somewhere. Filters apply first.
    /// </summary>
    public IReadOnlyList<AssetHit> Search(string? query, string? orientation, int minWidth, string? tag, int limit)
    {
        var words = Words(query ?? "");
        var hits = new List<AssetHit>();
        foreach (var entry in _entries.Values)
        {
            if (orientation is not null && !string.Equals(entry.Orientation, orientation, StringComparison.OrdinalIgnoreCase))
                continue;
            if (minWidth > 0 && entry.Width > 0 && entry.Width < minWidth)
                continue;
            if (tag is not null && !entry.Tags.Contains(tag.ToLowerInvariant()))
                continue;
            int score = 0;
            bool all = true;
            var name = Words(System.IO.Path.GetFileNameWithoutExtension(entry.Path));
            var folder = Words(System.IO.Path.GetDirectoryName(entry.Path) ?? "");
            var description = Words(entry.Description ?? "");
            foreach (var word in words)
            {
                int wordScore = (name.Any(item => item.StartsWith(word, StringComparison.Ordinal)) || entry.Tags.Any(item => item.StartsWith(word, StringComparison.Ordinal)) ? 3 : 0)
                    + (description.Any(item => item.StartsWith(word, StringComparison.Ordinal)) ? 2 : 0)
                    + (folder.Any(item => item.StartsWith(word, StringComparison.Ordinal)) ? 1 : 0);
                if (wordScore == 0)
                {
                    all = false;
                    break;
                }
                score += wordScore;
            }
            if (all)
                hits.Add(new AssetHit(entry, score));
        }
        return hits.OrderByDescending(hit => hit.Score).ThenByDescending(hit => (long)hit.Asset.Width * hit.Asset.Height)
            .ThenBy(hit => hit.Asset.Path, StringComparer.OrdinalIgnoreCase).Take(limit).ToList();
    }

    /// <summary>Groups of two or more entries with the same content hash.</summary>
    public IReadOnlyList<IReadOnlyList<AssetEntry>> Duplicates() =>
        _entries.Values.GroupBy(entry => entry.Hash).Where(group => group.Count() > 1)
            .Select(group => (IReadOnlyList<AssetEntry>)group.OrderBy(entry => entry.Path, StringComparer.OrdinalIgnoreCase).ToList())
            .OrderBy(group => group[0].Path, StringComparer.OrdinalIgnoreCase).ToList();

    /// <summary>Finds an entry by path.</summary>
    public AssetEntry? Find(string path) => _entries.GetValueOrDefault(System.IO.Path.GetFullPath(path));

    /// <summary>Adds a single file (used when tagging a file that was never scanned).</summary>
    public AssetEntry Add(string path)
    {
        var full = System.IO.Path.GetFullPath(path);
        var info = new FileInfo(full);
        if (!info.Exists)
            throw new ArgumentException($"The file '{full}' does not exist.");
        var entry = Read(full, info, info.LastWriteTimeUtc.ToString("o", CultureInfo.InvariantCulture))
            ?? throw new ArgumentException($"'{full}' is not a supported image.");
        _entries[full] = entry;
        return entry;
    }

    /// <summary>Lower-case words of a text (letters and digits, any script).</summary>
    public static List<string> Words(string text) =>
        [.. text.SplitWhere(static character => !char.IsLetterOrDigit(character)).Where(word => word.Length > 0).Select(word => word.ToLowerInvariant())];
}

internal static class StringSplitExtensions
{
    /// <summary>Splits on every character the predicate accepts.</summary>
    internal static string[] SplitWhere(this string text, Func<char, bool> isSeparator)
    {
        var parts = new List<string>();
        int start = 0;
        for (int index = 0; index <= text.Length; index++)
        {
            if (index == text.Length || isSeparator(text[index]))
            {
                if (index > start)
                    parts.Add(text[start..index]);
                start = index + 1;
            }
        }
        return [.. parts];
    }
}
