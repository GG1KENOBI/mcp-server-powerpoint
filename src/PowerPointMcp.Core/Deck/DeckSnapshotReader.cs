extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>
/// Reads inspection snapshots through typed PowerPoint COM members. Call only from inside
/// <c>IPresentationBatch.Execute</c> (the presentation's STA thread). Read-only: nothing is modified.
/// Every acquired COM object is released in a finally block.
/// </summary>
internal static class DeckSnapshotReader
{
    internal const int CompactTextLimit = 200;
    private const int CompactNotesLimit = 300;
    private const int MaxFontRunsRead = 200;

    /// <summary>Reads one slide's metadata (no objects).</summary>
    internal static DeckSlideInfo ReadSlide(
        PowerPoint.Slide slide,
        int slideIndex,
        IReadOnlyList<string> sectionNames,
        bool detailed,
        string? fingerprint = null)
    {
        PowerPoint.CustomLayout? layout = null;
        PowerPoint.SlideShowTransition? transition = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            layout = slide.CustomLayout;
            transition = slide.SlideShowTransition;
            shapes = slide.Shapes;
            var tags = ReadTags(slide.Tags);
            int section = sectionNames.Count > 0 ? slide.sectionIndex : 0;
            var notes = ReadNotes(slide);
            return new DeckSlideInfo
            {
                SlideIndex = slideIndex,
                SlideId = slide.SlideID,
                Name = slide.Name,
                Title = ReadTitle(shapes),
                LayoutName = layout.Name,
                Hidden = transition.Hidden == Office.MsoTriState.msoTrue,
                SectionIndex = section > 0 ? section : null,
                SectionName = section > 0 && section <= sectionNames.Count ? sectionNames[section - 1] : null,
                ShapeCount = shapes.Count,
                Notes = notes is null ? null : detailed ? notes : Shorten(notes, CompactNotesLimit, out _),
                AppId = tags.GetValueOrDefault(DeckRoles.IdTag),
                Tags = tags,
                Fingerprint = fingerprint,
            };
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (transition is not null) ComUtilities.Release(ref transition);
            if (layout is not null) ComUtilities.Release(ref layout);
        }
    }

    /// <summary>Reads every object on the slide, including members of groups (flattened, after their group).</summary>
    internal static List<DeckObjectInfo> ReadObjects(
        PowerPoint.Slide slide,
        int slideIndex,
        bool detailed,
        CancellationToken cancellationToken)
    {
        PowerPoint.Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            int slideId = slide.SlideID;
            var result = new List<DeckObjectInfo>(shapes.Count);
            for (int shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Shape? shape = null;
                try
                {
                    shape = shapes[shapeIndex];
                    ReadShapeTree(shape, slideIndex, slideId, shapeIndex, null, detailed, result);
                }
                finally
                {
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
            return result;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
        }
    }

    /// <summary>Names of the presentation's sections in order (empty when it has none).</summary>
    internal static List<string> ReadSectionNames(PowerPoint.Presentation presentation)
    {
        PowerPoint.SectionProperties? sections = null;
        try
        {
            sections = presentation.SectionProperties;
            var names = new List<string>(sections.Count);
            for (int index = 1; index <= sections.Count; index++)
            {
                names.Add(sections.Name(index));
            }
            return names;
        }
        finally
        {
            if (sections is not null) ComUtilities.Release(ref sections);
        }
    }

    /// <summary>Reads a slide with its objects and fingerprint.</summary>
    internal static (DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects) ReadSlideWithObjects(
        PowerPoint.Slide slide,
        int slideIndex,
        IReadOnlyList<string> sectionNames,
        bool detailed,
        CancellationToken cancellationToken)
    {
        // Fingerprints always hash full text so compact and detailed reads agree.
        var full = ReadObjects(slide, slideIndex, detailed: true, cancellationToken);
        var fingerprint = DeckFingerprint.ForSlide(slide.SlideID, full);
        var objects = detailed ? full : full.Select(Compact).ToList();
        return (ReadSlide(slide, slideIndex, sectionNames, detailed, fingerprint), objects);
    }

    /// <summary>Reads every slide (objects included) for fingerprints, validation, and diffs.</summary>
    internal static List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> ReadAll(
        PowerPoint.Presentation presentation,
        bool detailed,
        CancellationToken cancellationToken)
    {
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            var sectionNames = ReadSectionNames(presentation);
            var result = new List<(DeckSlideInfo, IReadOnlyList<DeckObjectInfo>)>(slides.Count);
            for (int index = 1; index <= slides.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Slide? slide = null;
                try
                {
                    slide = slides[index];
                    result.Add(ReadSlideWithObjects(slide, index, sectionNames, detailed, cancellationToken));
                }
                finally
                {
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
            return result;
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    /// <summary>Compact copy of an object: text shortened, detailed-only fields dropped.</summary>
    internal static DeckObjectInfo Compact(DeckObjectInfo item)
    {
        var text = item.Text is null ? null : Shorten(item.Text, CompactTextLimit, out _);
        return new DeckObjectInfo
        {
            SlideIndex = item.SlideIndex,
            SlideId = item.SlideId,
            ShapeIndex = item.ShapeIndex,
            ShapeId = item.ShapeId,
            Name = item.Name,
            Kind = item.Kind,
            Role = item.Role,
            PlaceholderType = item.PlaceholderType,
            AppId = item.AppId,
            Component = item.Component,
            GroupShapeId = item.GroupShapeId,
            MemberShapeIds = item.MemberShapeIds,
            Left = item.Left,
            Top = item.Top,
            Width = item.Width,
            Height = item.Height,
            Rotation = item.Rotation,
            ZOrder = item.ZOrder,
            Visible = item.Visible,
            HasText = item.HasText,
            Text = text,
            TextTruncated = item.Text is not null && text!.Length < item.Text.Length ? true : null,
            MinFontSize = item.MinFontSize,
            MaxFontSize = item.MaxFontSize,
            AutoSize = item.AutoSize,
            TextBounds = item.TextBounds,
            HasPicture = item.HasPicture,
            TableRows = item.TableRows,
            TableColumns = item.TableColumns,
            ChartType = item.ChartType,
            LinkSource = item.LinkSource,
            ConnectorBeginShapeId = item.ConnectorBeginShapeId,
            ConnectorEndShapeId = item.ConnectorEndShapeId,
        };
    }

    private static void ReadShapeTree(
        PowerPoint.Shape shape,
        int slideIndex,
        int slideId,
        int? shapeIndex,
        int? groupShapeId,
        bool detailed,
        List<DeckObjectInfo> result)
    {
        Office.MsoShapeType type = shape.Type;
        List<int>? memberIds = null;
        var members = new List<PowerPoint.Shape>();
        PowerPoint.GroupShapes? groupItems = null;
        try
        {
            if (type == Office.MsoShapeType.msoGroup)
            {
                groupItems = shape.GroupItems;
                memberIds = new List<int>(groupItems.Count);
                for (int memberIndex = 1; memberIndex <= groupItems.Count; memberIndex++)
                {
                    var member = groupItems[memberIndex];
                    members.Add(member);
                    memberIds.Add(member.Id);
                }
            }

            result.Add(ReadShape(shape, type, slideIndex, slideId, shapeIndex, groupShapeId, memberIds, detailed));
            foreach (var member in members)
            {
                ReadShapeTree(member, slideIndex, slideId, null, shape.Id, detailed, result);
            }
        }
        finally
        {
            for (int i = 0; i < members.Count; i++)
            {
                var member = members[i];
                ComUtilities.Release(ref member);
            }
            if (groupItems is not null) ComUtilities.Release(ref groupItems);
        }
    }

    private static DeckObjectInfo ReadShape(
        PowerPoint.Shape shape,
        Office.MsoShapeType type,
        int slideIndex,
        int slideId,
        int? shapeIndex,
        int? groupShapeId,
        IReadOnlyList<int>? memberIds,
        bool detailed)
    {
        bool isPlaceholder = type == Office.MsoShapeType.msoPlaceholder;
        bool isConnector = shape.Connector == Office.MsoTriState.msoTrue;
        string? placeholderType = isPlaceholder ? ReadPlaceholderType(shape) : null;
        var tags = ReadTags(shape.Tags);
        var text = ReadText(shape, detailed);
        var (rows, columns) = ReadTableSize(shape);
        var (begin, end) = isConnector ? ReadConnections(shape) : (null, null);
        return new DeckObjectInfo
        {
            SlideIndex = slideIndex,
            SlideId = slideId,
            ShapeIndex = shapeIndex,
            ShapeId = shape.Id,
            Name = shape.Name,
            Kind = KindOf(type, isConnector),
            Role = tags.GetValueOrDefault(DeckRoles.RoleTag) ?? DeckRoles.FromPlaceholder(placeholderType),
            PlaceholderType = placeholderType,
            AppId = tags.GetValueOrDefault(DeckRoles.IdTag),
            Component = tags.GetValueOrDefault(DeckRoles.ComponentTag),
            GroupShapeId = groupShapeId,
            MemberShapeIds = memberIds,
            Left = shape.Left,
            Top = shape.Top,
            Width = shape.Width,
            Height = shape.Height,
            Rotation = shape.Rotation,
            ZOrder = shapeIndex is null ? 0 : shape.ZOrderPosition,
            Visible = shape.Visible == Office.MsoTriState.msoTrue,
            HasText = text.Text is not null,
            Text = text.Text,
            ParagraphCount = text.ParagraphCount,
            MinFontSize = text.MinFontSize,
            MaxFontSize = text.MaxFontSize,
            FontNames = detailed ? text.FontNames : null,
            AutoSize = text.AutoSize,
            WordWrap = text.WordWrap,
            TextBounds = text.Bounds,
            TextMargins = detailed ? text.Margins : null,
            HasPicture = isPlaceholder ? HasPictureFill(shape) : null,
            TableRows = rows,
            TableColumns = columns,
            ChartType = ReadChartType(shape),
            LinkSource = type == Office.MsoShapeType.msoLinkedPicture ? ReadLinkSource(shape) : null,
            ConnectorBeginShapeId = begin,
            ConnectorEndShapeId = end,
            AltText = detailed ? shape.AlternativeText : null,
            Tags = detailed ? tags : null,
        };
    }

    internal static string KindOf(Office.MsoShapeType type, bool isConnector) => type switch
    {
        Office.MsoShapeType.msoPlaceholder => "placeholder",
        Office.MsoShapeType.msoTextBox => "text-box",
        Office.MsoShapeType.msoAutoShape when isConnector => "connector",
        Office.MsoShapeType.msoAutoShape or Office.MsoShapeType.msoCallout => "auto-shape",
        Office.MsoShapeType.msoLine => "line",
        Office.MsoShapeType.msoPicture or Office.MsoShapeType.msoLinkedPicture => "picture",
        Office.MsoShapeType.msoTable => "table",
        Office.MsoShapeType.msoChart => "chart",
        Office.MsoShapeType.msoSmartArt => "smart-art",
        Office.MsoShapeType.msoGroup => "group",
        Office.MsoShapeType.msoMedia => "media",
        Office.MsoShapeType.msoTextEffect => "text-effect",
        Office.MsoShapeType.msoFreeform => "freeform",
        _ => "other",
    };

    internal static Dictionary<string, string> ReadTags(PowerPoint.Tags tags)
    {
        try
        {
            var result = new Dictionary<string, string>(tags.Count, StringComparer.Ordinal);
            for (int index = 1; index <= tags.Count; index++)
            {
                result[tags.Name(index)] = tags.Value(index);
            }
            return result;
        }
        finally
        {
            ComUtilities.Release(ref tags!);
        }
    }

    private static string? ReadTitle(PowerPoint.Shapes shapes)
    {
        if (shapes.HasTitle != Office.MsoTriState.msoTrue)
            return null;

        PowerPoint.Shape? title = null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            title = shapes.Title;
            if (title.HasTextFrame != Office.MsoTriState.msoTrue)
                return null;
            frame = title.TextFrame;
            range = frame.TextRange;
            var text = range.Text;
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
            if (title is not null) ComUtilities.Release(ref title);
        }
    }

    private static string? ReadNotes(PowerPoint.Slide slide)
    {
        PowerPoint.SlideRange? notesPage = null;
        PowerPoint.Shapes? shapes = null;
        try
        {
            notesPage = slide.NotesPage;
            shapes = notesPage.Shapes;
            for (int index = 1; index <= shapes.Count; index++)
            {
                PowerPoint.Shape? shape = null;
                PowerPoint.PlaceholderFormat? format = null;
                PowerPoint.TextFrame? frame = null;
                PowerPoint.TextRange? range = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    if (format.Type != PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        continue;
                    frame = shape.TextFrame;
                    range = frame.TextRange;
                    var text = range.Text;
                    return string.IsNullOrWhiteSpace(text) ? null : text;
                }
                finally
                {
                    if (range is not null) ComUtilities.Release(ref range);
                    if (frame is not null) ComUtilities.Release(ref frame);
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
            return null;
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (notesPage is not null) ComUtilities.Release(ref notesPage);
        }
    }

    private static string ReadPlaceholderType(PowerPoint.Shape shape)
    {
        PowerPoint.PlaceholderFormat? format = null;
        try
        {
            format = shape.PlaceholderFormat;
            return format.Type.ToString();
        }
        finally
        {
            if (format is not null) ComUtilities.Release(ref format);
        }
    }

    private static bool HasPictureFill(PowerPoint.Shape shape)
    {
        PowerPoint.FillFormat? fill = null;
        try
        {
            fill = shape.Fill;
            return fill.Type == Office.MsoFillType.msoFillPicture;
        }
        finally
        {
            if (fill is not null) ComUtilities.Release(ref fill);
        }
    }

    private static (int? Rows, int? Columns) ReadTableSize(PowerPoint.Shape shape)
    {
        if (shape.HasTable != Office.MsoTriState.msoTrue)
            return (null, null);

        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        PowerPoint.Columns? columns = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            columns = table.Columns;
            return (rows.Count, columns.Count);
        }
        finally
        {
            if (columns is not null) ComUtilities.Release(ref columns);
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static int? ReadChartType(PowerPoint.Shape shape)
    {
        if (shape.HasChart != Office.MsoTriState.msoTrue)
            return null;

        PowerPoint.Chart? chart = null;
        try
        {
            chart = shape.Chart;
            return (int)chart.ChartType;
        }
        finally
        {
            if (chart is not null) ComUtilities.Release(ref chart);
        }
    }

    private static string? ReadLinkSource(PowerPoint.Shape shape)
    {
        PowerPoint.LinkFormat? link = null;
        try
        {
            link = shape.LinkFormat;
            return link.SourceFullName;
        }
        finally
        {
            if (link is not null) ComUtilities.Release(ref link);
        }
    }

    private static (int? Begin, int? End) ReadConnections(PowerPoint.Shape shape)
    {
        PowerPoint.ConnectorFormat? format = null;
        PowerPoint.Shape? begin = null;
        PowerPoint.Shape? end = null;
        try
        {
            format = shape.ConnectorFormat;
            int? beginId = null;
            int? endId = null;
            if (format.BeginConnected == Office.MsoTriState.msoTrue)
            {
                begin = format.BeginConnectedShape;
                beginId = begin.Id;
            }
            if (format.EndConnected == Office.MsoTriState.msoTrue)
            {
                end = format.EndConnectedShape;
                endId = end.Id;
            }
            return (beginId, endId);
        }
        finally
        {
            if (end is not null) ComUtilities.Release(ref end);
            if (begin is not null) ComUtilities.Release(ref begin);
            if (format is not null) ComUtilities.Release(ref format);
        }
    }

    private readonly record struct TextSnapshot(
        string? Text,
        int? ParagraphCount,
        float? MinFontSize,
        float? MaxFontSize,
        IReadOnlyList<string>? FontNames,
        string? AutoSize,
        bool? WordWrap,
        IReadOnlyList<float>? Bounds,
        IReadOnlyList<float>? Margins);

    private static TextSnapshot ReadText(PowerPoint.Shape shape, bool detailed)
    {
        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
            return default;

        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextFrame2? frame2 = null;
        PowerPoint.TextRange? range = null;
        PowerPoint.TextRange? paragraphs = null;
        try
        {
            frame = shape.TextFrame;
            if (frame.HasText != Office.MsoTriState.msoTrue)
                return default;

            range = frame.TextRange;
            string text = range.Text;
            if (string.IsNullOrWhiteSpace(text))
                return default;

            frame2 = shape.TextFrame2;
            paragraphs = range.Paragraphs();
            var (min, max, fonts) = ReadFonts(range, detailed);
            return new TextSnapshot(
                text,
                paragraphs.Count,
                min,
                max,
                fonts,
                AutoSizeName(frame2.AutoSize),
                frame.WordWrap == Office.MsoTriState.msoTrue,
                [range.BoundLeft, range.BoundTop, range.BoundWidth, range.BoundHeight],
                detailed ? [frame.MarginLeft, frame.MarginTop, frame.MarginRight, frame.MarginBottom] : null);
        }
        finally
        {
            if (paragraphs is not null) ComUtilities.Release(ref paragraphs);
            if (range is not null) ComUtilities.Release(ref range);
            if (frame2 is not null) ComUtilities.Release(ref frame2);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static (float? Min, float? Max, IReadOnlyList<string>? Fonts) ReadFonts(PowerPoint.TextRange range, bool detailed)
    {
        PowerPoint.TextRange? allRuns = null;
        float? min = null;
        float? max = null;
        var names = detailed ? new SortedSet<string>(StringComparer.Ordinal) : null;
        try
        {
            allRuns = range.Runs();
            int runCount = Math.Min(allRuns.Count, MaxFontRunsRead);
            for (int runIndex = 1; runIndex <= runCount; runIndex++)
            {
                PowerPoint.TextRange? run = null;
                PowerPoint.Font? font = null;
                try
                {
                    run = range.Runs(runIndex, 1);
                    font = run.Font;
                    float size = font.Size;
                    if (size > 0)
                    {
                        min = min is { } currentMin ? Math.Min(currentMin, size) : size;
                        max = max is { } currentMax ? Math.Max(currentMax, size) : size;
                    }
                    names?.Add(font.Name);
                }
                finally
                {
                    if (font is not null) ComUtilities.Release(ref font);
                    if (run is not null) ComUtilities.Release(ref run);
                }
            }
            return (min, max, names?.ToList());
        }
        finally
        {
            if (allRuns is not null) ComUtilities.Release(ref allRuns);
        }
    }

    private static string AutoSizeName(Office.MsoAutoSize autoSize) => autoSize switch
    {
        Office.MsoAutoSize.msoAutoSizeNone => "none",
        Office.MsoAutoSize.msoAutoSizeShapeToFitText => "shape-to-fit-text",
        Office.MsoAutoSize.msoAutoSizeTextToFitShape => "shrink-text-on-overflow",
        _ => "mixed",
    };

    internal static string Shorten(string value, int limit, out bool truncated)
    {
        truncated = value.Length > limit;
        return truncated ? value[..limit] : value;
    }
}
