using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.Core.Deck;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>The custom layouts compositions use, found by placeholder content (works in any UI language).</summary>
internal sealed record LayoutChoice(int? TitleOnlyIndex, Box? TitleBox, int? BlankIndex);

/// <summary>Finds the first master's "title only" and "blank" layouts by their placeholders. STA thread only.</summary>
internal static class LayoutPicker
{
    private static readonly HashSet<PowerPoint.PpPlaceholderType> Footers =
    [
        PowerPoint.PpPlaceholderType.ppPlaceholderDate,
        PowerPoint.PpPlaceholderType.ppPlaceholderFooter,
        PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber,
        PowerPoint.PpPlaceholderType.ppPlaceholderHeader,
    ];

    internal static LayoutChoice Pick(PowerPoint.Presentation presentation)
    {
        PowerPoint.Master? master = null;
        PowerPoint.CustomLayouts? layouts = null;
        int? titleOnly = null;
        int? blank = null;
        Box? titleBox = null;
        try
        {
            master = presentation.SlideMaster;
            layouts = master.CustomLayouts;
            for (int index = 1; index <= layouts.Count && (titleOnly is null || blank is null); index++)
            {
                PowerPoint.CustomLayout? layout = null;
                try
                {
                    layout = layouts[index];
                    var (types, title) = Placeholders(layout);
                    var content = types.Where(type => !Footers.Contains(type)).ToList();
                    if (titleOnly is null && content.Count == 1 && content[0] == PowerPoint.PpPlaceholderType.ppPlaceholderTitle && title is not null)
                    {
                        titleOnly = index;
                        titleBox = title;
                    }
                    else if (blank is null && content.Count == 0)
                    {
                        blank = index;
                    }
                }
                finally
                {
                    if (layout is not null) ComUtilities.Release(ref layout);
                }
            }
            return new LayoutChoice(titleOnly, titleBox, blank);
        }
        finally
        {
            if (layouts is not null) ComUtilities.Release(ref layouts);
            if (master is not null) ComUtilities.Release(ref master);
        }
    }

    private static (List<PowerPoint.PpPlaceholderType> Types, Box? Title) Placeholders(PowerPoint.CustomLayout layout)
    {
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Placeholders? placeholders = null;
        var types = new List<PowerPoint.PpPlaceholderType>();
        Box? title = null;
        try
        {
            shapes = layout.Shapes;
            placeholders = shapes.Placeholders;
            for (int index = 1; index <= placeholders.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                try
                {
                    shape = placeholders[index];
                    format = shape.PlaceholderFormat;
                    types.Add(format.Type);
                    if (format.Type == PowerPoint.PpPlaceholderType.ppPlaceholderTitle)
                        title = LayoutEngine.R(Box.FromSize(shape.Left, shape.Top, shape.Width, shape.Height));
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
            return (types, title);
        }
        finally
        {
            if (placeholders is not null) ComUtilities.Release(ref placeholders);
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    /// <summary>Adds a slide with the chosen custom layout, or the built-in layout type as a fallback.</summary>
    internal static PowerPoint.Slide AddSlide(PowerPoint.Presentation presentation, int index, int? layoutIndex, PowerPoint.PpSlideLayout fallback)
    {
        PowerPoint.Slides? slides = null;
        PowerPoint.Master? master = null;
        PowerPoint.CustomLayouts? layouts = null;
        PowerPoint.CustomLayout? layout = null;
        try
        {
            slides = presentation.Slides;
            var position = Math.Clamp(index, 1, slides.Count + 1);
            if (layoutIndex is { } chosen)
            {
                master = presentation.SlideMaster;
                layouts = master.CustomLayouts;
                layout = layouts[chosen];
                return slides.AddSlide(position, layout);
            }
            return slides.Add(position, fallback);
        }
        finally
        {
            if (layout is not null) ComUtilities.Release(ref layout);
            if (layouts is not null) ComUtilities.Release(ref layouts);
            if (master is not null) ComUtilities.Release(ref master);
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }
}
