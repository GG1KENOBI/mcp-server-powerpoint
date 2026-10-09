using System.Globalization;
using System.Text.RegularExpressions;
using Sbroenne.PowerPointMcp.Core.Deck;
using Sbroenne.PowerPointMcp.Core.Design;

namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>
/// Turns a validated composition into native-object placements. Pure and deterministic: the same
/// spec, profile, slide size, and context always produce the same plan. Text heights here are
/// estimates; the renderer measures with PowerPoint and fits.
/// </summary>
public static partial class CompositionPlanner
{
    private const float TakeawayMinHeight = 44f;

    /// <summary>Plans a composition.</summary>
    public static CompositionPlan Plan(CompositionSpec spec, ResolvedProfile profile, float width, float height, PlanContext context)
    {
        ArgumentNullException.ThrowIfNull(spec);
        ArgumentNullException.ThrowIfNull(profile);
        ArgumentNullException.ThrowIfNull(context);

        var planner = new Planner(spec, profile, width, height, context);
        var slides = spec.Kind switch
        {
            "title" => [planner.TitleSlide()],
            "section" => [planner.SectionSlide()],
            "executive-summary" => planner.ExecutiveSummary(),
            "comparison" => [planner.Comparison()],
            "cards" => [planner.Cards()],
            "kpis" => [planner.Kpis()],
            "image-text" => [planner.ImageText()],
            "table" => planner.Table(),
            "chart" => [planner.Chart()],
            "timeline" => [planner.Timeline()],
            "process" => [planner.Process()],
            "hierarchy" => [planner.Hierarchy()],
            "quote" => [planner.Quote()],
            "appendix" => planner.Appendix(),
            _ => throw new ArgumentException($"Unknown composition kind '{spec.Kind}'."),
        };

        return new CompositionPlan
        {
            Spec = spec,
            Slides = slides,
            ProfileName = profile.Name,
            Width = width,
            Height = height,
            ItemName = spec.Kind switch
            {
                "executive-summary" => "points",
                "table" => "rows",
                "appendix" => "references",
                _ => null,
            },
            Warnings = planner.Warnings,
        };
    }

    /// <summary>Continuation title: adds " (cont.)", or " (продолжение)" for Cyrillic titles.</summary>
    public static string ContinuationTitle(string title) =>
        Cyrillic().IsMatch(title) ? $"{title} (продолжение)" : $"{title} (cont.)";

    /// <summary>Whether a table cell looks numeric (amounts, percentages, signed values, currency).</summary>
    public static bool LooksNumeric(string cell) => NumericCell().IsMatch(cell.Trim());

    [GeneratedRegex(@"\p{IsCyrillic}", RegexOptions.CultureInvariant)]
    private static partial Regex Cyrillic();

    [GeneratedRegex(@"^[+\-−–(]?\s*[$€£₽¥]?\s*\d[\d\s.,  ]*\)?(\s*(%|‰|pts?|п\.п\.|x|×|[kKmMbB]|тыс\.?|млн\.?|млрд\.?|руб\.?|₽|\$|€))*$", RegexOptions.CultureInvariant)]
    private static partial Regex NumericCell();

    /// <summary>Shared planning state and helpers; kind layouts live in partial files.</summary>
    private sealed partial class Planner(CompositionSpec spec, ResolvedProfile profile, float width, float height, PlanContext context)
    {
        public List<string> Warnings { get; } = [];

        private CompositionSpec Spec { get; } = spec;

        private ResolvedProfile Profile { get; } = profile;

        private PlanContext Context { get; } = context;

        private float MinFont => Spec.Fit.MinFontSize ?? Profile.MinFontSize;

        private bool CanShrink => Spec.Fit.Policy.Contains("shrink");

        private bool CanSplit => Spec.Fit.Policy.Contains("split");

        private SlideFrame Frame(bool hasTitle = true) => ApplyTemplate(LayoutEngine.Frame(Profile, width, height, hasTitle, Spec.Source is not null), hasTitle);

        private SlideFrame ApplyTemplate(SlideFrame frame, bool hasTitle)
        {
            if (!hasTitle || Context.TemplateTitle is not { } title || title.Width <= 0 || title.Height <= 0)
                return frame;
            // Keep the deck's own title placement and start the body below it.
            var bodyTop = LayoutEngine.Round(Math.Max(title.Bottom + frame.Gap, frame.Safe.Top));
            var body = LayoutEngine.R(new Box(frame.Body.Left, Math.Min(bodyTop, frame.Body.Bottom), frame.Body.Right, frame.Body.Bottom));
            return frame with { Title = title, Body = body };
        }

        private string SlideId(int sequence) => sequence == 0 ? Context.AppId : $"{Context.AppId}~{sequence + 1}";

        private PlannedSlide Slide(int sequence, SlideBuilder builder, string? title, bool continuation = false, int offset = 0, int count = 0) => new()
        {
            Sequence = sequence,
            AppId = SlideId(sequence),
            Title = title,
            Continuation = continuation,
            ItemOffset = offset,
            ItemCount = count,
            Elements = builder.Elements,
            Regions = builder.Regions,
            Notes = sequence == 0 ? Spec.Notes : null,
        };

        private PlannedParagraph Text(string text, string sizeRole, string color = "text", bool bold = false, bool heading = false,
            string align = "left", string bullet = "none", int level = 0, float spaceAfter = 0, bool italic = false, float? size = null, string? hyperlink = null) => new()
            {
                Text = text,
                Size = size ?? Profile.Size(sizeRole),
                Font = heading ? Profile.HeadingFont : Profile.BodyFont,
                Color = Profile.Color(color),
                Bold = bold,
                Italic = italic,
                Align = align,
                Bullet = bullet,
                Level = level,
                SpaceAfter = spaceAfter,
                Hyperlink = hyperlink,
            };

        private static PlannedParagraph Scaled(PlannedParagraph paragraph, float factor) => new()
        {
            Text = paragraph.Text,
            Size = LayoutEngine.Round(Math.Max(1f, paragraph.Size * factor)),
            Font = paragraph.Font,
            Color = paragraph.Color,
            Bold = paragraph.Bold,
            Italic = paragraph.Italic,
            Align = paragraph.Align,
            Bullet = paragraph.Bullet,
            Level = paragraph.Level,
            SpaceAfter = paragraph.SpaceAfter * factor,
            Hyperlink = paragraph.Hyperlink,
        };

        /// <summary>Estimated height of paragraphs in a box of the given inner width.</summary>
        private static float Estimate(IReadOnlyList<PlannedParagraph> paragraphs, float innerWidth, float lineSpacing = 1f)
        {
            float total = 0;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                var indent = (paragraph.Bullet != "none" ? paragraph.Size * 1.1f : 0) + (paragraph.Level * 18f);
                total += TextEstimator.Height(paragraph.Text, paragraph.Size, innerWidth, TextEstimator.DefaultLineHeight * lineSpacing, 0, paragraph.Bold, indent);
                if (i < paragraphs.Count - 1)
                    total += paragraph.SpaceAfter;
            }
            return total;
        }

        private static float MinSize(IEnumerable<PlannedParagraph> paragraphs) => paragraphs.Select(paragraph => paragraph.Size).DefaultIfEmpty(0).Min();

        /// <summary>The largest factor (1 down to the floor, 2.5% steps) at which <paramref name="fits"/> holds; null when none.</summary>
        private float? LargestFittingFactor(float smallestSize, Func<float, bool> fits)
        {
            if (fits(1f))
                return 1f;
            if (!CanShrink || smallestSize <= 0)
                return null;
            var floor = Math.Min(1f, MinFont / smallestSize);
            for (var factor = 0.975f; factor >= floor - 0.0001f; factor -= 0.025f)
            {
                if (fits(factor))
                    return factor;
            }
            return null;
        }

        private PlannedElement TitleElement(SlideFrame frame, string title, string sizeRole = "title", string align = "left", string color = "text", Box? box = null, string vAlign = "bottom") => new()
        {
            Key = "title",
            Role = "title",
            Type = "text",
            Box = box ?? frame.Title,
            TitlePlaceholder = true,
            Paragraphs = [Text(title, sizeRole, color, heading: true, align: align)],
            VAlign = vAlign,
            Padding = [0, 2, 0, 2],
            Region = "title",
        };

        private void AddTitle(SlideBuilder builder, SlideFrame frame, string? title)
        {
            if (title is null)
                return;
            builder.Add(TitleElement(frame, title));
            builder.Region("title", frame.Title, "each", 0, ["title"], splittable: false, MinFont);
        }

        /// <summary>Adds the subtitle line at the top of the body and returns the remaining body.</summary>
        private Box AddSubtitle(SlideBuilder builder, Box body, string? subtitle)
        {
            if (subtitle is null)
                return body;
            var paragraphs = new[] { Text(subtitle, "subtitle", "muted") };
            var estimated = Estimate(paragraphs, body.Width);
            var (taken, rest) = LayoutEngine.TakeTop(body, MathF.Ceiling(estimated + 4f), Profile.Spacing("sm"));
            builder.Add(new PlannedElement
            {
                Key = "subtitle",
                Role = "subtitle",
                Type = "text",
                Box = taken,
                Paragraphs = paragraphs,
                Padding = [0, 0, 0, 0],
                Region = "subtitle",
                EstimatedTextHeight = estimated,
            });
            builder.Region("subtitle", taken, "each", 0, ["subtitle"], splittable: false, MinFont);
            return rest;
        }

        /// <summary>Adds the takeaway banner at the bottom of an area and returns the remaining area.</summary>
        private Box AddTakeaway(SlideBuilder builder, Box area, string? takeaway, bool vertical = false)
        {
            if (takeaway is null)
                return area;
            var callout = Profile.Component("callout");
            var paragraphs = new[] { Text(takeaway, "body", callout.Text, bold: true, heading: true, size: callout.BodySize) };
            var padding = callout.Padding;
            var estimated = Estimate(paragraphs, area.Width - (2 * padding) - 6);
            var bannerHeight = MathF.Ceiling(Math.Max(TakeawayMinHeight, estimated + (2 * padding)));
            var (banner, rest) = vertical ? (area, area) : LayoutEngine.TakeBottom(area, bannerHeight, Profile.Spacing("md"));
            builder.Add(new PlannedElement
            {
                Key = "takeaway",
                Role = "takeaway",
                Type = "shape",
                Geometry = callout.Radius > 0 ? "rounded-rectangle" : "rectangle",
                Radius = callout.Radius,
                Box = banner,
                Fill = callout.Fill,
                LineColor = callout.Border,
                LineWidth = callout.BorderWidth,
                Paragraphs = paragraphs,
                Padding = [padding + 6, padding, padding, padding],
                VAlign = vertical ? "top" : "middle",
                Component = callout.Identity,
                Region = "takeaway",
                EstimatedTextHeight = estimated,
            });
            builder.Add(new PlannedElement
            {
                Key = "takeaway.bar",
                Role = "decoration",
                Type = "shape",
                Box = LayoutEngine.R(Box.FromSize(banner.Left, banner.Top, 5, banner.Height)),
                Fill = callout.Accent,
                Component = callout.Identity,
                Group = "takeaway",
            });
            builder.Region("takeaway", banner, "each", 0, ["takeaway"], splittable: false, MinFont);
            return rest;
        }

        private void AddSource(SlideBuilder builder, SlideFrame frame)
        {
            if (Spec.Source is not { } source)
                return;
            var text = source.Contains(':', StringComparison.Ordinal) || source.StartsWith(Profile.Source.Prefix.Trim(), StringComparison.OrdinalIgnoreCase)
                ? source
                : Profile.Source.Prefix + source;
            builder.Add(new PlannedElement
            {
                Key = "source",
                Role = "source",
                Type = "text",
                Box = frame.Source,
                Paragraphs = [Text(text, "footnote", Profile.Source.Color, size: Profile.Source.FontSize)],
                VAlign = "bottom",
                Region = "source",
            });
            builder.Region("source", frame.Source, "each", 0, ["source"], splittable: false, Math.Min(MinFont, Profile.Source.FontSize));
        }

        private string ToneColor(string? tone, string fallback) => tone switch
        {
            "positive" => Profile.Color("positive"),
            "negative" => Profile.Color("negative"),
            "neutral" => Profile.Color("neutral"),
            _ => Profile.Color(fallback),
        };
    }

    /// <summary>Accumulates elements and regions for one slide.</summary>
    private sealed class SlideBuilder
    {
        private readonly HashSet<string> _keys = new(StringComparer.Ordinal);

        public List<PlannedElement> Elements { get; } = [];

        public List<FitRegion> Regions { get; } = [];

        public PlannedElement Add(PlannedElement element)
        {
            if (!_keys.Add(element.Key))
                throw new InvalidOperationException($"Duplicate element key '{element.Key}'.");
            Elements.Add(element);
            return element;
        }

        public void Region(string key, Box box, string mode, float gap, IReadOnlyList<string> keys, bool splittable, float minFontSize) =>
            Regions.Add(new FitRegion { Key = key, Box = box, Mode = mode, Gap = gap, Keys = keys, Splittable = splittable, MinFontSize = minFontSize });
    }

    private static string N(int value) => value.ToString(CultureInfo.InvariantCulture);
}
