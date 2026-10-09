extern alias OfficeInterop;

using Sbroenne.PowerPointMcp.ComInterop;
using Office = OfficeInterop::Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Where a block of searchable text lives.</summary>
internal sealed record DeckTextLocation(int SlideIndex, int SlideId, int ShapeId, string ShapeName, string? AppId, int? Row, int? Column, bool InNotes);

/// <summary>
/// Visits every block of editable text in a presentation: text frames (including group members
/// at any depth), every table cell, and optionally the speaker-notes body. Charts, SmartArt, and
/// OLE objects are not visited. The visitor receives the live TextRange, which is released after
/// the call. STA thread only.
/// </summary>
internal static class DeckTextWalker
{
    /// <param name="presentation">The presentation.</param>
    /// <param name="includeNotes">Visit the notes body placeholder of each slide.</param>
    /// <param name="includeShape">Filter on (SlideID, Shape.Id); a group that passes includes all of its members. Null visits everything.</param>
    /// <param name="includeNotesOf">Filter on SlideID for notes. Null visits every slide's notes.</param>
    /// <param name="visit">Called for each non-empty text range.</param>
    /// <param name="cancellationToken">Cancellation.</param>
    internal static void Visit(
        PowerPoint.Presentation presentation,
        bool includeNotes,
        Func<int, int, bool>? includeShape,
        Func<int, bool>? includeNotesOf,
        Action<DeckTextLocation, PowerPoint.TextRange> visit,
        CancellationToken cancellationToken)
    {
        PowerPoint.Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (int slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                PowerPoint.Slide? slide = null;
                PowerPoint.Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    int slideId = slide.SlideID;
                    shapes = slide.Shapes;
                    for (int index = 1; index <= shapes.Count; index++)
                    {
                        PowerPoint.Shape? shape = null;
                        try
                        {
                            shape = shapes[index];
                            VisitShape(shape, slideIndex, slideId, includeShape, visit, cancellationToken);
                        }
                        finally
                        {
                            if (shape is not null) ComUtilities.Release(ref shape);
                        }
                    }
                    if (includeNotes && (includeNotesOf is null || includeNotesOf(slideId)))
                        VisitNotes(slide, slideIndex, slideId, visit);
                }
                finally
                {
                    if (shapes is not null) ComUtilities.Release(ref shapes);
                    if (slide is not null) ComUtilities.Release(ref slide);
                }
            }
        }
        finally
        {
            if (slides is not null) ComUtilities.Release(ref slides);
        }
    }

    private static void VisitShape(
        PowerPoint.Shape shape,
        int slideIndex,
        int slideId,
        Func<int, int, bool>? includeShape,
        Action<DeckTextLocation, PowerPoint.TextRange> visit,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        int shapeId = shape.Id;
        bool included = includeShape is null || includeShape(slideId, shapeId);
        if (shape.Type == Office.MsoShapeType.msoGroup)
        {
            PowerPoint.GroupShapes? members = null;
            try
            {
                members = shape.GroupItems;
                for (int index = 1; index <= members.Count; index++)
                {
                    PowerPoint.Shape? member = null;
                    try
                    {
                        member = members[index];
                        VisitShape(member, slideIndex, slideId, included ? null : includeShape, visit, cancellationToken);
                    }
                    finally
                    {
                        if (member is not null) ComUtilities.Release(ref member);
                    }
                }
            }
            finally
            {
                if (members is not null) ComUtilities.Release(ref members);
            }
            return;
        }
        if (!included)
            return;

        string name = shape.Name;
        string? appId = ReadAppId(shape);
        if (shape.HasTable == Office.MsoTriState.msoTrue)
        {
            VisitTable(shape, new DeckTextLocation(slideIndex, slideId, shapeId, name, appId, null, null, false), visit, cancellationToken);
            return;
        }
        if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
            return;
        VisitFrame(shape, new DeckTextLocation(slideIndex, slideId, shapeId, name, appId, null, null, false), visit);
    }

    private static void VisitTable(PowerPoint.Shape shape, DeckTextLocation location, Action<DeckTextLocation, PowerPoint.TextRange> visit, CancellationToken cancellationToken)
    {
        PowerPoint.Table? table = null;
        PowerPoint.Rows? rows = null;
        PowerPoint.Columns? columns = null;
        try
        {
            table = shape.Table;
            rows = table.Rows;
            columns = table.Columns;
            int rowCount = rows.Count, columnCount = columns.Count;
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    PowerPoint.Cell? cell = null;
                    PowerPoint.Shape? cellShape = null;
                    try
                    {
                        cell = table.Cell(row, column);
                        cellShape = cell.Shape;
                        VisitFrame(cellShape, location with { Row = row, Column = column }, visit);
                    }
                    finally
                    {
                        if (cellShape is not null) ComUtilities.Release(ref cellShape);
                        if (cell is not null) ComUtilities.Release(ref cell);
                    }
                }
            }
        }
        finally
        {
            if (columns is not null) ComUtilities.Release(ref columns);
            if (rows is not null) ComUtilities.Release(ref rows);
            if (table is not null) ComUtilities.Release(ref table);
        }
    }

    private static void VisitNotes(PowerPoint.Slide slide, int slideIndex, int slideId, Action<DeckTextLocation, PowerPoint.TextRange> visit)
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
                try
                {
                    shape = shapes[index];
                    if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        continue;
                    format = shape.PlaceholderFormat;
                    if (format.Type != PowerPoint.PpPlaceholderType.ppPlaceholderBody)
                        continue;
                    VisitFrame(shape, new DeckTextLocation(slideIndex, slideId, shape.Id, shape.Name, null, null, null, true), visit);
                    return;
                }
                finally
                {
                    if (format is not null) ComUtilities.Release(ref format);
                    if (shape is not null) ComUtilities.Release(ref shape);
                }
            }
        }
        finally
        {
            if (shapes is not null) ComUtilities.Release(ref shapes);
            if (notesPage is not null) ComUtilities.Release(ref notesPage);
        }
    }

    private static void VisitFrame(PowerPoint.Shape shape, DeckTextLocation location, Action<DeckTextLocation, PowerPoint.TextRange> visit)
    {
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            frame = shape.TextFrame;
            if (frame.HasText != Office.MsoTriState.msoTrue)
                return;
            range = frame.TextRange;
            visit(location, range);
        }
        finally
        {
            if (range is not null) ComUtilities.Release(ref range);
            if (frame is not null) ComUtilities.Release(ref frame);
        }
    }

    private static string? ReadAppId(PowerPoint.Shape shape)
    {
        PowerPoint.Tags? tags = null;
        try
        {
            tags = shape.Tags;
            var value = tags[DeckRoles.IdTag];
            return string.IsNullOrEmpty(value) ? null : value;
        }
        finally
        {
            if (tags is not null) ComUtilities.Release(ref tags);
        }
    }
}
