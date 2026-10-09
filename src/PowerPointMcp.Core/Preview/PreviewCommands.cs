using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.PowerPointMcp.ComInterop;
using Sbroenne.PowerPointMcp.ComInterop.Session;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Review;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Sbroenne.PowerPointMcp.Core.Preview;

/// <summary>Preview commands.</summary>
public sealed class PreviewCommands : IPreviewCommands
{
    private const int CaptionHeight = 24;
    private const int Gap = 14;

    /// <inheritdoc/>
    public PreviewOperationResult Snapshot(
        IPresentationBatch batch,
        int? slideIndex = null,
        int? slideId = null,
        int width = 1280,
        string? outputPath = null,
        bool overwrite = false,
        bool includeFindings = true,
        bool includeSketch = true,
        bool refresh = false,
        bool includeImage = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (width is < 160 or > 3840)
            return Fail("width must be between 160 and 3840 pixels.");
        if (slideIndex is null && slideId is null)
            return Fail("Pass slide_index or slide_id.");
        if (CheckOutput(outputPath, overwrite) is { } outputError)
            return Fail(outputError);

        return batch.Execute((ctx, ct) =>
        {
            var deck = DeckCommands.ReadTargetSlides(ctx.Presentation, slideIndex, slideId, ct, out var error);
            if (error is not null)
                return Fail(error);
            var (slideInfo, objects) = deck[0];
            PowerPoint.PageSetup? setup = null;
            PowerPoint.Slide? slide = null;
            try
            {
                setup = ctx.Presentation.PageSetup;
                float slideWidth = setup.SlideWidth;
                float slideHeight = setup.SlideHeight;
                int height = (int)Math.Round(width * slideHeight / slideWidth);
                var cachePath = PreviewCache.PathFor(batch, slideInfo.SlideId, width, slideInfo.Fingerprint ?? "none");
                var target = outputPath ?? cachePath;
                bool cached = outputPath is null && !refresh && PreviewCache.IsFresh(cachePath);
                if (!cached)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                    slide = DeckShapeLocator.FindSlide(ctx.Presentation, slideInfo.SlideId)!;
                    try
                    {
                        slide.Export(target, "PNG", width, height);
                    }
                    catch (COMException ex)
                    {
                        return Fail($"PowerPoint could not export slide {slideInfo.SlideIndex} to '{target}': {ex.Message}. Check that the folder is writable and the path is not open elsewhere.");
                    }
                }

                IReadOnlyList<DeckFinding>? findings = null;
                if (includeFindings)
                {
                    findings = DeckValidator.Validate(slideWidth, slideHeight, deck, new ValidationOptions { InstalledFonts = FontCatalog.InstalledFamilies() }).Findings
                        .Where(finding => finding.SlideId == slideInfo.SlideId)
                        .ToList();
                }

                return new PreviewOperationResult
                {
                    Success = true,
                    ImagePath = target,
                    PixelWidth = width,
                    PixelHeight = height,
                    SlideWidth = slideWidth,
                    SlideHeight = slideHeight,
                    Scale = MathF.Round(width / slideWidth, 4),
                    SlideIndex = slideInfo.SlideIndex,
                    SlideId = slideInfo.SlideId,
                    Title = slideInfo.Title,
                    Fingerprint = slideInfo.Fingerprint,
                    Cached = cached,
                    Objects = objects.Select(DeckSnapshotReader.Compact).ToList(),
                    Findings = findings,
                    Sketch = includeSketch ? LayoutSketch.Render(slideWidth, slideHeight, objects) : null,
                    Warnings = cached ? ["Reused a cached render: the slide fingerprint and this session's change counter are unchanged. Pass refresh=true after editing the slide directly in PowerPoint."] : null,
                };
            }
            finally
            {
                if (slide is not null) ComUtilities.Release(ref slide);
                if (setup is not null) ComUtilities.Release(ref setup);
            }
        });
    }

    /// <inheritdoc/>
    public PreviewOperationResult ContactSheet(
        IPresentationBatch batch,
        int startSlide = 1,
        int maxSlides = 24,
        int columns = 4,
        int thumbnailWidth = 320,
        string? outputPath = null,
        bool overwrite = false,
        bool includeImage = true)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (maxSlides is < 1 or > 48)
            return Fail("max_slides must be between 1 and 48.");
        if (columns is < 1 or > 8)
            return Fail("columns must be between 1 and 8.");
        if (thumbnailWidth is < 120 or > 640)
            return Fail("thumbnail_width must be between 120 and 640 pixels.");
        if (startSlide < 1)
            return Fail("start_slide must be 1 or greater.");
        if (CheckOutput(outputPath, overwrite) is { } outputError)
            return Fail(outputError);

        var work = Path.Combine(PreviewCache.Directory, "sheet-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(work);
        try
        {
            var rendered = batch.Execute((ctx, ct) =>
            {
                PowerPoint.PageSetup? setup = null;
                PowerPoint.Slides? slides = null;
                try
                {
                    setup = ctx.Presentation.PageSetup;
                    slides = ctx.Presentation.Slides;
                    int count = slides.Count;
                    if (startSlide > Math.Max(1, count))
                        return (Error: $"start_slide {startSlide} is beyond the last slide ({count}).", Items: new List<(int, int, string?, bool, string)>(), Height: 0);
                    int height = (int)Math.Round(thumbnailWidth * setup.SlideHeight / setup.SlideWidth);
                    var items = new List<(int Index, int Id, string? Title, bool Hidden, string Path)>();
                    for (int index = startSlide; index <= Math.Min(count, startSlide + maxSlides - 1); index++)
                    {
                        ct.ThrowIfCancellationRequested();
                        PowerPoint.Slide? slide = null;
                        try
                        {
                            slide = slides[index];
                            var info = DeckSnapshotReader.ReadSlide(slide, index, [], detailed: false);
                            var path = Path.Combine(work, $"{index.ToString(CultureInfo.InvariantCulture)}.png");
                            slide.Export(path, "PNG", thumbnailWidth, height);
                            items.Add((index, info.SlideId, info.Title, info.Hidden, path));
                        }
                        finally
                        {
                            if (slide is not null) ComUtilities.Release(ref slide);
                        }
                    }
                    return (Error: (string?)null, Items: items, Height: height);
                }
                finally
                {
                    if (slides is not null) ComUtilities.Release(ref slides);
                    if (setup is not null) ComUtilities.Release(ref setup);
                }
            });
            if (rendered.Error is not null)
                return Fail(rendered.Error);
            if (rendered.Items.Count == 0)
                return Fail("The presentation has no slides.");

            var target = outputPath ?? Path.Combine(PreviewCache.Directory, $"contact-sheet-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..6]}.png");
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            int rows = (rendered.Items.Count + columns - 1) / columns;
            int sheetWidth = Gap + (columns * (thumbnailWidth + Gap));
            int sheetHeight = Gap + (rows * (rendered.Height + CaptionHeight + Gap));
            var placed = new List<ContactSheetSlide>();
            using (var sheet = new Bitmap(sheetWidth, sheetHeight, PixelFormat.Format24bppRgb))
            using (var graphics = Graphics.FromImage(sheet))
            using (var font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Pixel))
            using (var bold = new Font("Segoe UI", 10f, FontStyle.Bold, GraphicsUnit.Pixel))
            using (var border = new Pen(Color.FromArgb(200, 200, 200)))
            using (var hiddenBrush = new SolidBrush(Color.FromArgb(160, 255, 255, 255)))
            {
                graphics.Clear(Color.White);
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                for (int i = 0; i < rendered.Items.Count; i++)
                {
                    var (index, id, title, hidden, path) = rendered.Items[i];
                    int column = i % columns;
                    int row = i / columns;
                    int x = Gap + (column * (thumbnailWidth + Gap));
                    int y = Gap + (row * (rendered.Height + CaptionHeight + Gap));
                    using (var thumbnail = System.Drawing.Image.FromFile(path))
                        graphics.DrawImage(thumbnail, x, y, thumbnailWidth, rendered.Height);
                    graphics.DrawRectangle(border, x - 1, y - 1, thumbnailWidth + 1, rendered.Height + 1);
                    if (hidden)
                    {
                        graphics.FillRectangle(hiddenBrush, x, y, thumbnailWidth, rendered.Height);
                        graphics.DrawString("hidden", bold, Brushes.DimGray, x + 6, y + 6);
                    }
                    var number = index.ToString(CultureInfo.InvariantCulture);
                    graphics.DrawString(number, bold, Brushes.Black, x, y + rendered.Height + 4);
                    var numberWidth = (int)Math.Ceiling(graphics.MeasureString(number + " ", bold).Width);
                    var caption = new RectangleF(x + numberWidth, y + rendered.Height + 4, thumbnailWidth - numberWidth, CaptionHeight - 4);
                    using var format = new StringFormat { Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap };
                    graphics.DrawString((title ?? "(no title)").Replace('\r', ' ').Replace('\n', ' ').Replace('\v', ' '), font, Brushes.DimGray, caption, format);
                    placed.Add(new ContactSheetSlide(index, id, title, hidden, column + 1, row + 1));
                }
                sheet.Save(target, ImageFormat.Png);
            }

            return new PreviewOperationResult
            {
                Success = true,
                ImagePath = target,
                PixelWidth = sheetWidth,
                PixelHeight = sheetHeight,
                Slides = placed,
            };
        }
        finally
        {
            try
            {
                Directory.Delete(work, recursive: true);
            }
            catch (IOException)
            {
                // Thumbnails are temporary; a locked file is cleaned up by clear-cache later.
            }
        }
    }

    /// <inheritdoc/>
    public PreviewOperationResult CompareImages(IPresentationBatch batch, string firstPath, string secondPath, int threshold = 24)
    {
        ArgumentNullException.ThrowIfNull(batch);
        if (threshold is < 0 or > 255)
            return Fail("threshold must be between 0 and 255.");
        foreach (var path in new[] { firstPath, secondPath })
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return Fail($"'{path}' does not exist.");
        }

        using var first = new Bitmap(firstPath);
        using var second = new Bitmap(secondPath);
        if (first.Width != second.Width || first.Height != second.Height)
            return Fail($"The images differ in size ({first.Width}x{first.Height} vs {second.Width}x{second.Height}); render both at the same width.");
        var a = Pixels(first);
        var b = Pixels(second);
        var difference = ImageDiff.Compare(first.Width, first.Height, first.Width * 4, a, b, threshold);

        // Convert the changed box to slide points using the current slide size.
        var size = batch.Execute((ctx, ct) =>
        {
            PowerPoint.PageSetup? setup = null;
            try
            {
                setup = ctx.Presentation.PageSetup;
                return (setup.SlideWidth, setup.SlideHeight);
            }
            finally
            {
                if (setup is not null) ComUtilities.Release(ref setup);
            }
        });
        IReadOnlyList<float>? area = null;
        if (difference.Left is { } left)
        {
            var scale = size.SlideWidth / first.Width;
            area = [MathF.Round(left * scale, 1), MathF.Round(difference.Top!.Value * scale, 1),
                MathF.Round((difference.Right!.Value - left + 1) * scale, 1), MathF.Round((difference.Bottom!.Value - difference.Top.Value + 1) * scale, 1)];
        }
        return new PreviewOperationResult { Success = true, Difference = difference, ChangedArea = area, SlideWidth = size.SlideWidth, SlideHeight = size.SlideHeight };
    }

    /// <inheritdoc/>
    public PreviewOperationResult ClearCache(IPresentationBatch batch, bool all = false) =>
        new() { Success = true, Removed = PreviewCache.Prune(all) };

    private static byte[] Pixels(Bitmap bitmap)
    {
        var rectangle = new Rectangle(0, 0, bitmap.Width, bitmap.Height);
        var data = bitmap.LockBits(rectangle, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            var buffer = new byte[bitmap.Width * bitmap.Height * 4];
            for (int y = 0; y < bitmap.Height; y++)
                Marshal.Copy(data.Scan0 + (y * data.Stride), buffer, y * bitmap.Width * 4, bitmap.Width * 4);
            return buffer;
        }
        finally
        {
            bitmap.UnlockBits(data);
        }
    }

    private static string? CheckOutput(string? outputPath, bool overwrite)
    {
        if (outputPath is null)
            return null;
        if (!Path.IsPathFullyQualified(outputPath))
            return "output_path must be a full path (e.g. C:\\Users\\me\\Documents\\slide3.png).";
        if (!outputPath.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
            return "output_path must end in .png.";
        if (File.Exists(outputPath) && !overwrite)
            return $"'{outputPath}' already exists. Pass overwrite=true to replace it.";
        return null;
    }

    private static PreviewOperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };
}
