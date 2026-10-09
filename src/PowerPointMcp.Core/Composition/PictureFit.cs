using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>How a picture fills its frame.</summary>
public sealed record PicturePlacement(
    Box Frame,
    float PictureWidth,
    float PictureHeight,
    float OffsetX,
    float OffsetY,
    bool Cropped);

/// <summary>
/// Contain / cover / crop-to-focal-point math for placing a picture into a frame without
/// distortion. Pure. Offsets follow PowerPoint's Crop object: the picture center relative to
/// the frame center, in points.
/// </summary>
public static class PictureFit
{
    /// <summary>
    /// contain: the whole picture is visible, centered, frame shrinks to the picture.
    /// cover: the picture fills the frame and the overflow is cropped around the focal point
    /// (0..1 from the left/top; 0.5, 0.5 is the center). stretch is never used.
    /// </summary>
    public static PicturePlacement Compute(Box frame, double sourceAspect, string fit, float focalX = 0.5f, float focalY = 0.5f)
    {
        if (sourceAspect <= 0 || frame.Width <= 0 || frame.Height <= 0)
            throw new ArgumentException("Frame and source aspect must be positive.");
        var frameAspect = frame.Width / (double)frame.Height;
        focalX = Math.Clamp(focalX, 0f, 1f);
        focalY = Math.Clamp(focalY, 0f, 1f);

        if (fit == "contain")
        {
            float width = frameAspect > sourceAspect ? (float)(frame.Height * sourceAspect) : frame.Width;
            float height = frameAspect > sourceAspect ? frame.Height : (float)(frame.Width / sourceAspect);
            var placed = LayoutEngine.Center(frame, width, height);
            return new PicturePlacement(placed, placed.Width, placed.Height, 0, 0, false);
        }

        if (fit != "cover")
            throw new ArgumentException($"Unknown fit '{fit}'. Use contain or cover.", nameof(fit));

        float pictureWidth = frameAspect > sourceAspect ? frame.Width : (float)(frame.Height * sourceAspect);
        float pictureHeight = frameAspect > sourceAspect ? (float)(frame.Width / sourceAspect) : frame.Height;
        // Offset so the focal point stays visible: shift the picture center toward the frame center.
        var maxOffsetX = (pictureWidth - frame.Width) / 2f;
        var maxOffsetY = (pictureHeight - frame.Height) / 2f;
        var offsetX = Math.Clamp((0.5f - focalX) * pictureWidth, -maxOffsetX, maxOffsetX);
        var offsetY = Math.Clamp((0.5f - focalY) * pictureHeight, -maxOffsetY, maxOffsetY);
        bool cropped = maxOffsetX > 0.05f || maxOffsetY > 0.05f;
        return new PicturePlacement(frame, LayoutEngine.Round(pictureWidth), LayoutEngine.Round(pictureHeight),
            LayoutEngine.Round(offsetX), LayoutEngine.Round(offsetY), cropped);
    }
}
