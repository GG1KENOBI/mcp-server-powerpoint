using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>Resolution, distortion, alt-text, and link checks for a picture. Pure.</summary>
public static class PictureQuality
{
    /// <summary>Below this many pixels per inch a picture looks soft when projected.</summary>
    public const int SoftPpi = 150;

    /// <summary>Below this many pixels per inch a picture looks blurry.</summary>
    public const int LowPpi = 96;

    /// <summary>Builds the picture report from the snapshot and the uncropped picture size. Pure.</summary>
    public static DeckPictureInfo Describe(DeckObjectInfo item, float pictureWidth, float pictureHeight)
    {
        var tags = item.Tags ?? new Dictionary<string, string>();
        bool linked = item.LinkSource is not null;
        var pixels = item.SourcePixelSize;
        bool cropped = item.Crop is { } crop && crop.Any(value => MathF.Abs(value) > 0.5f);
        int? ppi = pixels is { Count: 2 } && pixels[0] > 0 && pictureWidth > 0 ? (int)Math.Round(pixels[0] / (pictureWidth / 72.0)) : null;
        float? distortion = null;
        if (pixels is { Count: 2 } && pixels[0] > 0 && pixels[1] > 0 && pictureWidth > 0 && pictureHeight > 0)
        {
            var expected = pixels[0] / (double)pixels[1];
            var shown = pictureWidth / (double)pictureHeight;
            distortion = (float)Math.Round(Math.Abs(shown / expected - 1) * 1000) / 10f;
        }
        bool decorative = tags.GetValueOrDefault(AssetCommands.DecorativeTag) == "1";
        var issues = new List<string>();
        if (!decorative && string.IsNullOrWhiteSpace(item.AltText))
            issues.Add("missing-alt-text");
        if (ppi is < LowPpi)
            issues.Add("low-resolution");
        else if (ppi is < SoftPpi)
            issues.Add("soft-resolution");
        if (ppi is null)
            issues.Add("unknown-resolution");
        if (distortion is > 2f)
            issues.Add("distorted");
        if (linked && item.LinkSourceExists == false)
            issues.Add("broken-link");
        else if (linked)
            issues.Add("linked");
        return new DeckPictureInfo
        {
            SlideIndex = item.SlideIndex,
            SlideId = item.SlideId,
            ShapeId = item.ShapeId,
            Name = item.Name,
            AppId = item.AppId,
            Linked = linked,
            LinkSource = item.LinkSource,
            LinkExists = linked ? item.LinkSourceExists : null,
            Source = tags.GetValueOrDefault(AssetCommands.SourceTag),
            PixelSize = pixels,
            DisplaySize = [MathF.Round(item.Width * 10f) / 10f, MathF.Round(item.Height * 10f) / 10f],
            EffectivePpi = ppi,
            Cropped = cropped,
            DistortionPercent = distortion,
            AltText = string.IsNullOrWhiteSpace(item.AltText) ? null : item.AltText,
            Decorative = decorative,
            Attribution = tags.GetValueOrDefault(DeckRoles.AttributionTag),
            Issues = issues.Count > 0 ? issues : null,
        };
    }
}
