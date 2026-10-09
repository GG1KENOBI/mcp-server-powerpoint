extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Finds slides by SlideID and shapes by Shape.Id (searching inside groups). STA thread only.</summary>
internal static class DeckShapeLocator
{
    /// <summary>Returns the slide with the SlideID, or null. The caller releases it.</summary>
    internal static PowerPoint.Slide? FindSlide(PowerPoint.Presentation presentation, int slideId)
    {
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int index = 1; index <= slides.Count; index++)
            {
                var slide = slides[index];
                if (slide.SlideID == slideId)
                    return slide;
                ComUtilities.Release(ref slide!);
            }
            return null;
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    /// <summary>Returns the shape with the Shape.Id on the slide (top level or inside a group), or null. The caller releases it.</summary>
    internal static PowerPoint.Shape? FindShape(PowerPoint.Slide slide, int shapeId)
    {
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                var found = Search(shape, shapeId);
                if (!ReferenceEquals(found, shape))
                    ComUtilities.Release(ref shape!);
                if (found is not null)
                    return found;
            }
            return null;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    private static PowerPoint.Shape? Search(PowerPoint.Shape shape, int shapeId)
    {
        if (shape.Id == shapeId)
            return shape;
        if (shape.Type != Office.MsoShapeType.msoGroup)
            return null;

        PowerPoint.GroupShapes? members = null;
        try
        {
            members = shape.GroupItems;
            for (int index = 1; index <= members.Count; index++)
            {
                var member = members[index];
                var found = Search(member, shapeId);
                if (!ReferenceEquals(found, member))
                    ComUtilities.Release(ref member!);
                if (found is not null)
                    return found;
            }
            return null;
        }
        finally
        {
            if (members is not null) ComUtilities.Release(ref members);
        }
    }
}
