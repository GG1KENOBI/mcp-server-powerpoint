namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>Result of comparing two images of the same size.</summary>
public sealed record ImageDiffResult(int Width, int Height, long ChangedPixels, double ChangedRatio, int? Left, int? Top, int? Right, int? Bottom);

/// <summary>Pixel comparison of two 32-bit BGRA buffers. Pure.</summary>
public static class ImageDiff
{
    /// <summary>
    /// Counts pixels whose largest channel difference exceeds <paramref name="threshold"/> (0-255)
    /// and returns the bounding box of the changed pixels.
    /// </summary>
    public static ImageDiffResult Compare(int width, int height, int stride, ReadOnlySpan<byte> first, ReadOnlySpan<byte> second, int threshold = 24)
    {
        long changed = 0;
        int left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1;
        for (int y = 0; y < height; y++)
        {
            int row = y * stride;
            for (int x = 0; x < width; x++)
            {
                int offset = row + (x * 4);
                int difference = Math.Max(Math.Abs(first[offset] - second[offset]),
                    Math.Max(Math.Abs(first[offset + 1] - second[offset + 1]), Math.Abs(first[offset + 2] - second[offset + 2])));
                if (difference <= threshold)
                    continue;
                changed++;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x);
                bottom = Math.Max(bottom, y);
            }
        }
        return changed == 0
            ? new ImageDiffResult(width, height, 0, 0, null, null, null, null)
            : new ImageDiffResult(width, height, changed, changed / (double)((long)width * height), left, top, right, bottom);
    }
}
