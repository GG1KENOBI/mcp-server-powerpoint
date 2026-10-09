extern alias OfficeInterop;

using System.Globalization;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Composition;
using Sbroenne.PowerPointMcp.Core.Deck;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Assets;

/// <summary>Places local pictures without distortion (contain, or cover with a focal crop). STA thread only.</summary>
internal static class PicturePlacer
{
    /// <summary>Adds an embedded picture (never linked) placed and cropped per the plan, tagged with its pixel size, source, and attribution.</summary>
    internal static PowerPoint.Shape Place(PowerPoint.Shapes shapes, PicturePlan picture)
    {
        var placement = picture.Placement;
        if (picture.Fit == "contain" || !placement.Cropped)
        {
            var frame = placement.Frame;
            var shape = shapes.AddPicture(picture.Path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, frame.Left, frame.Top, frame.Width, frame.Height);
            TagPixels(shape, picture);
            return shape;
        }

        var cover = shapes.AddPicture(picture.Path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, placement.Frame.Left, placement.Frame.Top, placement.PictureWidth, placement.PictureHeight);
        PowerPoint.PictureFormat? format = null;
        Office.Crop? crop = null;
        try
        {
            format = cover.PictureFormat;
            crop = format.Crop;
            crop.PictureWidth = placement.PictureWidth;
            crop.PictureHeight = placement.PictureHeight;
            crop.ShapeLeft = placement.Frame.Left;
            crop.ShapeTop = placement.Frame.Top;
            crop.ShapeWidth = placement.Frame.Width;
            crop.ShapeHeight = placement.Frame.Height;
            crop.PictureOffsetX = placement.OffsetX;
            crop.PictureOffsetY = placement.OffsetY;
        }
        finally
        {
            if (crop is not null) ComUtilities.Release(ref crop);
            if (format is not null) ComUtilities.Release(ref format);
        }
        TagPixels(cover, picture);
        return cover;
    }

    internal static void TagPixels(PowerPoint.Shape shape, PicturePlan picture)
    {
        var tags = shape.Tags;
        try
        {
            tags.Add(DeckRoles.ImagePixelsTag, $"{picture.PixelWidth.ToString(CultureInfo.InvariantCulture)}x{picture.PixelHeight.ToString(CultureInfo.InvariantCulture)}");
            if (picture.Attribution is not null)
                tags.Add(DeckRoles.AttributionTag, picture.Attribution);
            tags.Add("PPTMCP_IMG_SOURCE", picture.Path);
        }
        finally
        {
            ComUtilities.Release(ref tags);
        }
    }
}
