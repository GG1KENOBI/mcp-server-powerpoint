namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>A bullet or numbered point: headline text with optional detail and nesting level (0-4).</summary>
public sealed record SpecPoint(string Text, string? Detail = null, int Level = 0);

/// <summary>A column of a comparison.</summary>
public sealed record SpecColumn(string Heading, IReadOnlyList<SpecPoint> Points, string? Tone = null);

/// <summary>A card.</summary>
public sealed record SpecCard(string Heading, string? Body, string? Badge);

/// <summary>
/// A key performance indicator. Trend is the arrow direction (up, down, flat); sentiment is the
/// color (positive, negative, neutral) and defaults to up=positive, down=negative, flat=neutral,
/// so a rising churn rate is {"trend":"up","sentiment":"negative"}.
/// </summary>
public sealed record SpecKpi(string Value, string Label, string? Delta, string? Trend, string? Note, string? Sentiment = null);

/// <summary>A local image.</summary>
public sealed record SpecImage(string Path, string? Alt, string Fit, float FocalX, float FocalY, string Position, string? Attribution);

/// <summary>Table content. Cells are text exactly as given; numbers are not reformatted.</summary>
public sealed record SpecTable(
    IReadOnlyList<string> Header,
    IReadOnlyList<IReadOnlyList<string>> Rows,
    IReadOnlyList<string>? Align,
    bool TotalRow,
    IReadOnlyList<int>? Highlight,
    IReadOnlyList<float>? ColumnWeights,
    bool RepeatHeader);

/// <summary>A chart series.</summary>
public sealed record SpecSeries(string Name, IReadOnlyList<double> Values);

/// <summary>Chart content.</summary>
public sealed record SpecChart(
    string Type,
    IReadOnlyList<string> Categories,
    IReadOnlyList<SpecSeries> Series,
    string? NumberFormat,
    string? ValueAxisTitle,
    string? CategoryAxisTitle,
    bool? Legend,
    bool? DataLabels);

/// <summary>A timeline milestone.</summary>
public sealed record SpecMilestone(string Date, string Label, string? Detail, string Status);

/// <summary>A process step.</summary>
public sealed record SpecStep(string Label, string? Detail);

/// <summary>A node of a hierarchy.</summary>
public sealed record SpecNode(string Label, string? Detail, IReadOnlyList<SpecNode> Children, string? Id = null);

/// <summary>A quote.</summary>
public sealed record SpecQuote(string Text, string Author, string? Role);

/// <summary>A reference in an appendix.</summary>
public sealed record SpecReference(string Text, string? Label, string? Url);

/// <summary>Fitting policy for text that does not fit.</summary>
public sealed record SpecFit(IReadOnlyList<string> Policy, float? MinFontSize);

/// <summary>
/// A validated composition (schema <c>pptmcp.composition/1</c>). Content fields are kind-specific;
/// <see cref="CompositionKinds"/> documents which ones each kind uses.
/// </summary>
public sealed record CompositionSpec
{
    /// <summary>Schema identifier.</summary>
    public const string SchemaId = "pptmcp.composition/1";

    /// <summary>Composition kind, e.g. executive-summary.</summary>
    public required string Kind { get; init; }

    /// <summary>Persistent id for the slide (PPTMCP_ID); generated when omitted.</summary>
    public string? Id { get; init; }

    /// <summary>Slide title.</summary>
    public string? Title { get; init; }

    /// <summary>Subtitle or kicker line.</summary>
    public string? Subtitle { get; init; }

    /// <summary>1-based position to insert at (default: after the last slide).</summary>
    public int? InsertAt { get; init; }

    /// <summary>Design profile name (default "theme": the deck's own theme colors and fonts).</summary>
    public string? Profile { get; init; }

    /// <summary>Speaker notes.</summary>
    public string? Notes { get; init; }

    /// <summary>Source / citation line.</summary>
    public string? Source { get; init; }

    /// <summary>Key message shown as a highlighted takeaway.</summary>
    public string? Takeaway { get; init; }

    /// <summary>Date line (title kind).</summary>
    public string? Date { get; init; }

    /// <summary>Author / presenter line (title kind).</summary>
    public string? Author { get; init; }

    /// <summary>Section number (section kind).</summary>
    public string? SectionNumber { get; init; }

    /// <summary>Points (executive-summary, image-text, appendix).</summary>
    public IReadOnlyList<SpecPoint>? Points { get; init; }

    /// <summary>Columns (comparison: exactly two).</summary>
    public IReadOnlyList<SpecColumn>? Columns { get; init; }

    /// <summary>Cards.</summary>
    public IReadOnlyList<SpecCard>? Cards { get; init; }

    /// <summary>KPIs.</summary>
    public IReadOnlyList<SpecKpi>? Kpis { get; init; }

    /// <summary>Image (image-text).</summary>
    public SpecImage? Image { get; init; }

    /// <summary>Table.</summary>
    public SpecTable? Table { get; init; }

    /// <summary>Chart.</summary>
    public SpecChart? Chart { get; init; }

    /// <summary>Milestones (timeline).</summary>
    public IReadOnlyList<SpecMilestone>? Milestones { get; init; }

    /// <summary>Steps (process).</summary>
    public IReadOnlyList<SpecStep>? Steps { get; init; }

    /// <summary>Root node (hierarchy).</summary>
    public SpecNode? Tree { get; init; }

    /// <summary>Quote.</summary>
    public SpecQuote? Quote { get; init; }

    /// <summary>References (appendix).</summary>
    public IReadOnlyList<SpecReference>? References { get; init; }

    /// <summary>Fitting policy.</summary>
    public SpecFit Fit { get; init; } = new(["shrink", "split"], null);

    /// <summary>Free-form provenance stored as PPTMCP_META_* slide tags (source_ref, generated_by, ...).</summary>
    public IReadOnlyDictionary<string, string>? Metadata { get; init; }

    /// <summary>The original JSON, stored on the slide so it can be re-rendered or updated.</summary>
    public string? OriginalJson { get; init; }
}
