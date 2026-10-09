using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>Format and pixel size of an image file, read from its header.</summary>
public sealed class ImageHeaderInfo
{
    /// <summary>png, jpeg, gif, bmp, webp, tiff, svg, emf, or wmf.</summary>
    public required string Format { get; init; }

    /// <summary>Width in pixels (for SVG: user units; for EMF: device units of the frame).</summary>
    public required int Width { get; init; }

    /// <summary>Height in pixels (for SVG: user units; for EMF: device units of the frame).</summary>
    public required int Height { get; init; }

    /// <summary>EXIF orientation (1-8) for JPEG files that declare one; 5-8 mean the image is shown rotated 90 degrees.</summary>
    public int? ExifOrientation { get; init; }

    /// <summary>Whether the format is vector (svg, emf, wmf).</summary>
    public bool IsVector => Format is "svg" or "emf" or "wmf";

    /// <summary>Displayed width after EXIF rotation.</summary>
    public int DisplayWidth => ExifOrientation is >= 5 and <= 8 ? Height : Width;

    /// <summary>Displayed height after EXIF rotation.</summary>
    public int DisplayHeight => ExifOrientation is >= 5 and <= 8 ? Width : Height;

    /// <summary>Displayed width divided by displayed height.</summary>
    public double Aspect => DisplayHeight == 0 ? 0 : (double)DisplayWidth / DisplayHeight;
}

/// <summary>
/// Reads image format and dimensions from file headers without decoding pixels and without
/// any platform imaging library. Pure and offline; unknown or corrupt files return null.
/// </summary>
public static partial class ImageHeaderReader
{
    private const int MaxHeaderBytes = 512 * 1024;

    /// <summary>File extensions this reader understands.</summary>
    public static readonly IReadOnlySet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        ".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".dib", ".webp", ".tif", ".tiff", ".svg", ".emf", ".wmf",
    };

    /// <summary>Reads the header of a file; returns null when the file is missing, unreadable, or not a supported image.</summary>
    public static ImageHeaderInfo? ReadFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return null;
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var length = (int)Math.Min(stream.Length, MaxHeaderBytes);
            var buffer = new byte[length];
            stream.ReadExactly(buffer, 0, length);
            return Read(buffer);
        }
        catch (IOException)
        {
            return null;
        }
        catch (UnauthorizedAccessException)
        {
            return null;
        }
    }

    /// <summary>Reads image info from the first bytes of a file (up to 512 KB is enough for every supported format).</summary>
    public static ImageHeaderInfo? Read(ReadOnlySpan<byte> data)
    {
        try
        {
            if (data.Length >= 24 && data[..8].SequenceEqual(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }))
                return Info("png", BinaryPrimitives.ReadInt32BigEndian(data[16..]), BinaryPrimitives.ReadInt32BigEndian(data[20..]));
            if (data.Length >= 4 && data[0] == 0xFF && data[1] == 0xD8)
                return ReadJpeg(data);
            if (data.Length >= 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F')
                return Info("gif", BinaryPrimitives.ReadUInt16LittleEndian(data[6..]), BinaryPrimitives.ReadUInt16LittleEndian(data[8..]));
            if (data.Length >= 26 && data[0] == 'B' && data[1] == 'M')
                return Info("bmp", BinaryPrimitives.ReadInt32LittleEndian(data[18..]), Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(data[22..])));
            if (data.Length >= 30 && data[..4].SequenceEqual("RIFF"u8) && data[8..12].SequenceEqual("WEBP"u8))
                return ReadWebP(data);
            if (data.Length >= 8 && ((data[0] == 'I' && data[1] == 'I' && data[2] == 42) || (data[0] == 'M' && data[1] == 'M' && data[3] == 42)))
                return ReadTiff(data);
            if (data.Length >= 44 && BinaryPrimitives.ReadUInt32LittleEndian(data) == 1 && data[40..44].SequenceEqual(" EMF"u8))
                return ReadEmf(data);
            if (data.Length >= 22 && BinaryPrimitives.ReadUInt32LittleEndian(data) == 0x9AC6CDD7)
                return ReadPlaceableWmf(data);
            return ReadSvg(data);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    private static ImageHeaderInfo? Info(string format, int width, int height, int? orientation = null) =>
        width > 0 && height > 0 ? new ImageHeaderInfo { Format = format, Width = width, Height = height, ExifOrientation = orientation } : null;

    private static ImageHeaderInfo? ReadJpeg(ReadOnlySpan<byte> data)
    {
        int? orientation = null;
        int offset = 2;
        while (offset + 4 <= data.Length)
        {
            if (data[offset] != 0xFF)
            {
                offset++;
                continue;
            }
            byte marker = data[offset + 1];
            if (marker == 0xFF)
            {
                offset++;
                continue;
            }
            if (marker is 0xD8 or 0x01 || (marker >= 0xD0 && marker <= 0xD7))
            {
                offset += 2;
                continue;
            }
            int length = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 2)..]);
            if (length < 2)
                return null;
            if (marker == 0xE1 && orientation is null)
                orientation = ReadExifOrientation(data.Slice(offset + 4, Math.Min(length - 2, data.Length - offset - 4)));
            // SOF0-SOF15 except DHT (C4), JPG (C8), DAC (CC).
            if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
            {
                int height = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 5)..]);
                int width = BinaryPrimitives.ReadUInt16BigEndian(data[(offset + 7)..]);
                return Info("jpeg", width, height, orientation);
            }
            if (marker == 0xDA)
                return null;
            offset += 2 + length;
        }
        return null;
    }

    private static int? ReadExifOrientation(ReadOnlySpan<byte> app1)
    {
        if (app1.Length < 14 || !app1[..6].SequenceEqual("Exif\0\0"u8))
            return null;
        var tiff = app1[6..];
        bool little = tiff[0] == 'I';
        int ifd = (int)ReadUInt32(tiff[4..], little);
        if (ifd < 8 || ifd + 2 > tiff.Length)
            return null;
        int entries = ReadUInt16(tiff[ifd..], little);
        for (int i = 0; i < entries; i++)
        {
            int entry = ifd + 2 + (i * 12);
            if (entry + 12 > tiff.Length)
                return null;
            if (ReadUInt16(tiff[entry..], little) == 0x0112)
            {
                int value = ReadUInt16(tiff[(entry + 8)..], little);
                return value is >= 1 and <= 8 ? value : null;
            }
        }
        return null;
    }

    private static ImageHeaderInfo? ReadWebP(ReadOnlySpan<byte> data)
    {
        var chunk = data[12..16];
        if (chunk.SequenceEqual("VP8 "u8) && data.Length >= 30)
            return Info("webp", BinaryPrimitives.ReadUInt16LittleEndian(data[26..]) & 0x3FFF, BinaryPrimitives.ReadUInt16LittleEndian(data[28..]) & 0x3FFF);
        if (chunk.SequenceEqual("VP8L"u8) && data.Length >= 25)
        {
            uint bits = BinaryPrimitives.ReadUInt32LittleEndian(data[21..]);
            return Info("webp", (int)(bits & 0x3FFF) + 1, (int)((bits >> 14) & 0x3FFF) + 1);
        }
        if (chunk.SequenceEqual("VP8X"u8) && data.Length >= 30)
        {
            int width = 1 + (data[24] | (data[25] << 8) | (data[26] << 16));
            int height = 1 + (data[27] | (data[28] << 8) | (data[29] << 16));
            return Info("webp", width, height);
        }
        return null;
    }

    private static ImageHeaderInfo? ReadTiff(ReadOnlySpan<byte> data)
    {
        bool little = data[0] == 'I';
        int ifd = (int)ReadUInt32(data[4..], little);
        if (ifd < 8 || ifd + 2 > data.Length)
            return null;
        int entries = ReadUInt16(data[ifd..], little);
        int width = 0, height = 0;
        for (int i = 0; i < entries; i++)
        {
            int entry = ifd + 2 + (i * 12);
            if (entry + 12 > data.Length)
                break;
            int tag = ReadUInt16(data[entry..], little);
            int type = ReadUInt16(data[(entry + 2)..], little);
            int value = type == 3 ? ReadUInt16(data[(entry + 8)..], little) : (int)ReadUInt32(data[(entry + 8)..], little);
            if (tag == 256) width = value;
            if (tag == 257) height = value;
        }
        return Info("tiff", width, height);
    }

    private static ImageHeaderInfo? ReadEmf(ReadOnlySpan<byte> data)
    {
        // EMR_HEADER rclFrame (0.01 mm) at offset 24; use the device bounds at offset 8 for pixels.
        int left = BinaryPrimitives.ReadInt32LittleEndian(data[8..]);
        int top = BinaryPrimitives.ReadInt32LittleEndian(data[12..]);
        int right = BinaryPrimitives.ReadInt32LittleEndian(data[16..]);
        int bottom = BinaryPrimitives.ReadInt32LittleEndian(data[20..]);
        return Info("emf", right - left + 1, bottom - top + 1);
    }

    private static ImageHeaderInfo? ReadPlaceableWmf(ReadOnlySpan<byte> data)
    {
        int left = BinaryPrimitives.ReadInt16LittleEndian(data[6..]);
        int top = BinaryPrimitives.ReadInt16LittleEndian(data[8..]);
        int right = BinaryPrimitives.ReadInt16LittleEndian(data[10..]);
        int bottom = BinaryPrimitives.ReadInt16LittleEndian(data[12..]);
        int unitsPerInch = BinaryPrimitives.ReadUInt16LittleEndian(data[14..]);
        if (unitsPerInch <= 0)
            return null;
        // Express as 96-dpi pixels.
        return Info("wmf", (int)Math.Round((right - left) * 96.0 / unitsPerInch), (int)Math.Round((bottom - top) * 96.0 / unitsPerInch));
    }

    private static ImageHeaderInfo? ReadSvg(ReadOnlySpan<byte> data)
    {
        var text = Encoding.UTF8.GetString(data[..Math.Min(data.Length, 8192)]);
        var match = SvgRoot().Match(text);
        if (!match.Success)
            return null;
        var tag = match.Value;
        double? width = SvgLength(Attribute(tag, "width"));
        double? height = SvgLength(Attribute(tag, "height"));
        var viewBox = Attribute(tag, "viewBox")?.Split([' ', ','], StringSplitOptions.RemoveEmptyEntries);
        if ((width is null || height is null) && viewBox is { Length: 4 } &&
            double.TryParse(viewBox[2], NumberStyles.Float, CultureInfo.InvariantCulture, out var boxWidth) &&
            double.TryParse(viewBox[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var boxHeight))
        {
            if (width is null && height is null)
                (width, height) = (boxWidth, boxHeight);
            else if (width is null && boxHeight > 0)
                width = height!.Value * boxWidth / boxHeight;
            else if (boxWidth > 0)
                height = width!.Value * boxHeight / boxWidth;
        }
        return width is > 0 && height is > 0 ? Info("svg", (int)Math.Round(width.Value), (int)Math.Round(height.Value)) : null;
    }

    private static string? Attribute(string tag, string name)
    {
        var match = Regex.Match(tag, $@"\s{name}\s*=\s*[""']([^""']*)[""']", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
        return match.Success ? match.Groups[1].Value : null;
    }

    private static double? SvgLength(string? value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.EndsWith('%'))
            return null;
        var match = SvgNumber().Match(value.Trim());
        if (!match.Success || !double.TryParse(match.Groups["n"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return null;
        return match.Groups["u"].Value switch
        {
            "" or "px" => number,
            "pt" => number * 96 / 72,
            "in" => number * 96,
            "cm" => number * 96 / 2.54,
            "mm" => number * 96 / 25.4,
            _ => null,
        };
    }

    private static int ReadUInt16(ReadOnlySpan<byte> data, bool little) =>
        little ? BinaryPrimitives.ReadUInt16LittleEndian(data) : BinaryPrimitives.ReadUInt16BigEndian(data);

    private static uint ReadUInt32(ReadOnlySpan<byte> data, bool little) =>
        little ? BinaryPrimitives.ReadUInt32LittleEndian(data) : BinaryPrimitives.ReadUInt32BigEndian(data);

    [GeneratedRegex(@"<svg\b[^>]*>", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)]
    private static partial Regex SvgRoot();

    [GeneratedRegex(@"^(?<n>[0-9]*\.?[0-9]+)(?<u>px|pt|in|cm|mm)?$", RegexOptions.CultureInvariant)]
    private static partial Regex SvgNumber();
}
