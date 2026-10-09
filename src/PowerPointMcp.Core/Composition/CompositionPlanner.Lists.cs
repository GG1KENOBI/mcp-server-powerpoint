using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Composition;

public static partial class CompositionPlanner
{
    private sealed partial class Planner
    {
        private const float NumberSize = 28f;

        public List<PlannedSlide> ExecutiveSummary()
        {
            var points = Spec.Points!;
            var gap = Profile.Spacing("sm");
            var probe = Frame();
            var bodyWidth = probe.Body.Width - NumberSize - Profile.Spacing("md");

            IReadOnlyList<PlannedParagraph> PointText(SpecPoint point, float factor)
            {
                var list = new List<PlannedParagraph> { Scaled(Text(point.Text, "heading", bold: true, heading: true, spaceAfter: 2), factor) };
                if (point.Detail is not null)
                    list.Add(Scaled(Text(point.Detail, "body", "muted"), factor));
                return list;
            }

            float ItemHeight(int index, float factor) => MathF.Ceiling(Estimate(PointText(points[index], factor), bodyWidth, Profile.LineSpacing) + 6f);

            var smallest = MinSize(points.SelectMany(point => PointText(point, 1f)));
            var (partition, factor) = PartitionItems(points.Count, ItemHeight, gap, smallest, hasTakeaway: Spec.Takeaway is not null);

            var slides = new List<PlannedSlide>();
            int offset = 0;
            for (int sequence = 0; sequence < partition.Count; sequence++)
            {
                var builder = new SlideBuilder();
                var frame = Frame();
                var title = sequence == 0 ? Spec.Title! : ContinuationTitle(Spec.Title!);
                AddTitle(builder, frame, title);
                var body = sequence == 0 ? AddSubtitle(builder, frame.Body, Spec.Subtitle) : frame.Body;
                body = sequence == partition.Count - 1 ? AddTakeaway(builder, body, Spec.Takeaway) : body;

                var keys = new List<string>();
                var heights = Enumerable.Range(offset, partition[sequence]).Select(index => ItemHeight(index, factor)).ToList();
                var boxes = LayoutEngine.Stack(body, heights, gap);
                for (int i = 0; i < partition[sequence]; i++)
                {
                    int number = offset + i + 1;
                    var key = $"point-{N(number)}";
                    var box = boxes[i];
                    var textBox = LayoutEngine.R(new Box(box.Left + NumberSize + Profile.Spacing("md"), box.Top, box.Right, box.Bottom));
                    var accent = Profile.Component("badge");
                    builder.Add(new PlannedElement
                    {
                        Key = $"{key}.number",
                        Role = "point-number",
                        Type = "shape",
                        Geometry = "oval",
                        Box = LayoutEngine.R(Box.FromSize(box.Left, box.Top + 2, NumberSize, NumberSize)),
                        Fill = Profile.Color("primary"),
                        Paragraphs = [Text(N(number), "caption", "background", bold: true, heading: true, align: "center", size: 12)],
                        VAlign = "middle",
                        AnchorKey = key,
                        AnchorOffset = 2,
                        Component = accent.Identity,
                    });
                    builder.Add(new PlannedElement
                    {
                        Key = key,
                        Role = "point",
                        Type = "text",
                        Box = textBox,
                        Paragraphs = PointText(points[offset + i], factor),
                        LineSpacing = Profile.LineSpacing,
                        Padding = [0, 3, 0, 3],
                        Region = "points",
                        EstimatedTextHeight = heights[i],
                    });
                    keys.Add(key);
                }
                builder.Region("points", body, "stack", gap, keys, splittable: CanSplit, MinFont);
                AddSource(builder, frame);
                slides.Add(Slide(sequence, builder, title, sequence > 0, offset, partition[sequence]));
                offset += partition[sequence];
            }
            return slides;
        }

        /// <summary>
        /// Decides how many items go on each slide and the font factor: everything on one slide at
        /// full size, else shrunk (if allowed), else split at full size (if allowed), else one
        /// slide at the smallest allowed size with an overflow warning.
        /// </summary>
        private (List<int> Partition, float Factor) PartitionItems(int count, Func<int, float, float> itemHeight, float gap, float smallestSize, bool hasTakeaway)
        {
            var frame = Frame();
            var subtitleHeight = Spec.Subtitle is null ? 0 : Estimate([Text(Spec.Subtitle, "subtitle")], frame.Body.Width) + 4 + Profile.Spacing("sm");
            var takeawayHeight = hasTakeaway ? TakeawayMinHeight + Profile.Spacing("md") : 0;
            float Available(bool first, bool last) => frame.Body.Height - (first ? subtitleHeight : 0) - (last ? takeawayHeight : 0);

            if (Context.Partition is { } forced)
            {
                if (forced.Sum() != count || forced.Any(value => value < 1))
                    throw new ArgumentException("The forced partition does not cover every item.");
                return ([.. forced], 1f);
            }

            float Total(int start, int length, float factor) =>
                LayoutEngine.StackHeight(Enumerable.Range(start, length).Select(index => itemHeight(index, factor)).ToList(), gap);

            var single = Available(first: true, last: true);
            var factor = LargestFittingFactor(smallestSize, f => Total(0, count, f) <= single);
            if (factor is { } fitting)
                return ([count], fitting);

            if (CanSplit && count > 1)
            {
                var partition = new List<int>();
                int start = 0;
                while (start < count)
                {
                    bool first = partition.Count == 0;
                    int take = 0;
                    while (start + take < count && Total(start, take + 1, 1f) <= Available(first, last: start + take + 1 == count))
                        take++;
                    if (take == 0)
                    {
                        Warnings.Add($"Item {start + 1} alone is estimated taller than a slide's body; it is placed on its own slide and may overflow.");
                        take = 1;
                    }
                    partition.Add(take);
                    start += take;
                }
                Warnings.Add($"Content is estimated not to fit one slide; planned {partition.Count} slides ({string.Join(" + ", partition)} items). Rendering measures and adjusts.");
                return (partition, 1f);
            }

            var floor = CanShrink && smallestSize > 0 ? Math.Min(1f, MinFont / smallestSize) : 1f;
            Warnings.Add($"Content is estimated to overflow even at {MinFont:0.#} pt and splitting is not allowed by fit.policy; overflow will be reported, never truncated.");
            return ([count], floor);
        }

        public PlannedSlide Comparison()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            body = AddTakeaway(builder, body, Spec.Takeaway);
            var columns = LayoutEngine.Columns(body, 2, Profile.Gutter * 1.5f);
            var card = Profile.Component("card");
            var headerHeight = MathF.Ceiling((Profile.Size("heading") * 1.3f) + 16f);

            IReadOnlyList<PlannedParagraph> Bullets(SpecColumn column, float factor) =>
                column.Points.Select(point => Scaled(Text(point.Detail is null ? point.Text : $"{point.Text}: {point.Detail}", "body",
                    bullet: "bullet", level: point.Level, spaceAfter: Profile.ParagraphSpacing), factor)).ToList();

            var panelInner = columns[0].Width - (2 * card.Padding);
            var panelHeight = body.Height - headerHeight;
            var smallest = MinSize(Spec.Columns!.SelectMany(column => Bullets(column, 1f)));
            var factor = LargestFittingFactor(smallest, f => Spec.Columns!.All(column => Estimate(Bullets(column, f), panelInner, Profile.LineSpacing) + (2 * card.Padding) <= panelHeight));
            if (factor is null)
                Warnings.Add("The comparison columns are estimated to overflow at the minimum font size; shorten the points or split into two slides.");

            var keys = new List<string>();
            for (int i = 0; i < 2; i++)
            {
                var column = Spec.Columns![i];
                var (header, panel) = LayoutEngine.TakeTop(columns[i], headerHeight, 0);
                var key = $"column-{N(i + 1)}";
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.heading",
                    Role = "column-heading",
                    Type = "shape",
                    Box = header,
                    Fill = ToneColor(column.Tone, "primary"),
                    Paragraphs = [Text(column.Heading, "heading", "background", bold: true, heading: true)],
                    VAlign = "middle",
                    Padding = [card.Padding, 4, card.Padding, 4],
                    Group = key,
                    Component = card.Identity,
                });
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.body",
                    Role = "column-body",
                    Type = "shape",
                    Box = panel,
                    Fill = card.Fill,
                    LineColor = card.Border,
                    LineWidth = card.BorderWidth,
                    Paragraphs = Bullets(column, factor ?? Math.Min(1f, MinFont / Math.Max(1f, smallest))),
                    LineSpacing = Profile.LineSpacing,
                    Padding = [card.Padding, card.Padding, card.Padding, card.Padding],
                    Group = key,
                    Component = card.Identity,
                    Region = "columns",
                });
                keys.Add($"{key}.body");
            }
            builder.Region("columns", body, "each", 0, keys, splittable: false, MinFont);
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide Cards()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            body = AddTakeaway(builder, body, Spec.Takeaway);
            var cards = Spec.Cards!;
            var style = Profile.Component("card");
            var badge = Profile.Component("badge");
            var gutter = Profile.Gutter;
            int columns = LayoutEngine.AutoColumns(cards.Count, body.Width, gutter, 180f);
            int rows = (cards.Count + columns - 1) / columns;
            var cellWidth = (body.Width - (gutter * (columns - 1))) / columns;
            var inner = cellWidth - (2 * style.Padding);

            PlannedParagraph Heading(SpecCard card, float factor) => Scaled(Text(card.Heading, "heading", style.Text, bold: true, heading: true, size: style.HeadingSize), factor);
            IReadOnlyList<PlannedParagraph> Body(SpecCard card, float factor) =>
                card.Body is null ? [] : [Scaled(Text(card.Body, "body", "muted", size: style.BodySize), factor)];
            float Needed(SpecCard card, float factor) =>
                Estimate([Heading(card, factor)], inner - (card.Badge is null ? 0 : 50)) + Profile.Spacing("sm") +
                Estimate(Body(card, factor), inner, Profile.LineSpacing) + (2 * style.Padding) + 6;

            var available = (body.Height - (gutter * (rows - 1))) / rows;
            var smallest = MinSize(cards.SelectMany(card => Body(card, 1f).Append(Heading(card, 1f))));
            var factor = LargestFittingFactor(smallest, f => cards.All(card => Needed(card, f) <= available));
            if (factor is null)
                Warnings.Add("Some cards are estimated to overflow at the minimum font size; shorten them or use fewer cards.");
            var f = factor ?? Math.Min(1f, MinFont / Math.Max(1f, smallest));
            var cardHeight = Math.Min(available, Math.Max(110f, MathF.Ceiling(cards.Max(card => Needed(card, f)) + 8)));
            var area = new Box(body.Left, body.Top, body.Right, body.Top + (cardHeight * rows) + (gutter * (rows - 1)));
            var boxes = LayoutEngine.Grid(area, cards.Count, columns, gutter, gutter);

            var keys = new List<string>();
            for (int i = 0; i < cards.Count; i++)
            {
                var card = cards[i];
                var box = boxes[i];
                var key = $"card-{N(i + 1)}";
                var content = LayoutEngine.Inset(box, style.Padding, style.Padding + 4, style.Padding, style.Padding);
                var heading = Heading(card, f);
                var headingWidth = content.Width - (card.Badge is null ? 0 : 56);
                var headingHeight = MathF.Ceiling(Estimate([heading], headingWidth) + 4);
                var (headingBox, bodyBox) = LayoutEngine.TakeTop(content, headingHeight, Profile.Spacing("sm"));
                headingBox = LayoutEngine.R(new Box(headingBox.Left, headingBox.Top, headingBox.Left + headingWidth, headingBox.Bottom));

                builder.Add(new PlannedElement
                {
                    Key = $"{key}.background",
                    Role = "card",
                    Type = "shape",
                    Geometry = style.Radius > 0 ? "rounded-rectangle" : "rectangle",
                    Radius = style.Radius,
                    Box = box,
                    Fill = style.Fill,
                    LineColor = style.Border,
                    LineWidth = style.BorderWidth,
                    Group = key,
                    Component = style.Identity,
                });
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.bar",
                    Role = "decoration",
                    Type = "shape",
                    Box = LayoutEngine.R(Box.FromSize(box.Left + style.Radius, box.Top, box.Width - (2 * style.Radius), 4)),
                    Fill = style.Accent,
                    Group = key,
                    Component = style.Identity,
                });
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.heading",
                    Role = "card-heading",
                    Type = "text",
                    Box = headingBox,
                    Paragraphs = [heading],
                    Group = key,
                    Component = style.Identity,
                    Region = "cards",
                });
                keys.Add($"{key}.heading");
                if (card.Body is not null)
                {
                    builder.Add(new PlannedElement
                    {
                        Key = $"{key}.body",
                        Role = "card-body",
                        Type = "text",
                        Box = bodyBox,
                        Paragraphs = Body(card, f),
                        LineSpacing = Profile.LineSpacing,
                        Group = key,
                        Component = style.Identity,
                        Region = "cards",
                    });
                    keys.Add($"{key}.body");
                }
                if (card.Badge is not null)
                {
                    var badgeText = Text(card.Badge, "caption", badge.Text, bold: true, align: "center", size: badge.BodySize);
                    var badgeWidth = MathF.Ceiling(Math.Min(content.Width / 2f, TextEstimator.LineWidth(card.Badge, badge.BodySize, bold: true) + 16));
                    builder.Add(new PlannedElement
                    {
                        Key = $"{key}.badge",
                        Role = "badge",
                        Type = "shape",
                        Geometry = "rounded-rectangle",
                        Radius = badge.Radius,
                        Box = LayoutEngine.R(Box.FromSize(content.Right - badgeWidth, content.Top, badgeWidth, badge.BodySize + 10)),
                        Fill = badge.Fill,
                        Paragraphs = [badgeText],
                        VAlign = "middle",
                        Padding = [4, 1, 4, 1],
                        Group = key,
                        Component = badge.Identity,
                    });
                }
            }
            builder.Region("cards", body, "each", 0, keys, splittable: false, MinFont);
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide Kpis()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            body = AddTakeaway(builder, body, Spec.Takeaway);
            var kpis = Spec.Kpis!;
            var style = Profile.Component("kpi");
            int columns = kpis.Count <= 4 ? kpis.Count : kpis.Count <= 6 ? 3 : 4;
            int rows = (kpis.Count + columns - 1) / columns;
            var gutter = Profile.Gutter;
            var cellWidth = (body.Width - (gutter * (columns - 1))) / columns;
            var inner = cellWidth - (2 * style.Padding) - 6;
            var valueSize = style.HeadingSize;
            var labelSize = style.BodySize;
            var deltaSize = Profile.Size("body");

            float Needed(SpecKpi kpi, float factor) =>
                Estimate([Scaled(Text(kpi.Value, "kpi_value", size: valueSize, bold: true, heading: true), factor)], inner) +
                Estimate([Scaled(Text(kpi.Label, "kpi_label", size: labelSize), factor)], inner) + 4 +
                (kpi.Delta is null ? 0 : Estimate([Scaled(Text(kpi.Delta, "body", size: deltaSize, bold: true), factor)], inner - 16) + 4) +
                (kpi.Note is null ? 0 : Estimate([Scaled(Text(kpi.Note, "caption"), factor)], inner) + 4) + (2 * style.Padding);

            var available = (body.Height - (gutter * (rows - 1))) / rows;
            var smallest = Math.Min(labelSize, Profile.Size("caption"));
            var factor = LargestFittingFactor(Math.Max(smallest, MinFont), f => kpis.All(kpi => Needed(kpi, f) <= available)) ?? 1f;
            if (kpis.Any(kpi => Needed(kpi, factor) > available))
                Warnings.Add("Some KPI panels are estimated to overflow; shorten labels or notes.");
            var panelHeight = Math.Min(available, Math.Max(120f, MathF.Ceiling(kpis.Max(kpi => Needed(kpi, factor)) + 8)));
            var area = new Box(body.Left, body.Top, body.Right, body.Top + (panelHeight * rows) + (gutter * (rows - 1)));
            var boxes = LayoutEngine.Grid(area, kpis.Count, columns, gutter, gutter);

            var keys = new List<string>();
            for (int i = 0; i < kpis.Count; i++)
            {
                var kpi = kpis[i];
                var box = boxes[i];
                var key = $"kpi-{N(i + 1)}";
                var sentiment = kpi.Sentiment ?? kpi.Trend switch
                {
                    "up" => "positive",
                    "down" => "negative",
                    "flat" => "neutral",
                    _ => null,
                };
                var toneColor = ToneColor(sentiment, "primary");
                var content = LayoutEngine.Inset(box, style.Padding + 6, style.Padding, style.Padding, style.Padding);

                builder.Add(new PlannedElement
                {
                    Key = $"{key}.background",
                    Role = "kpi",
                    Type = "shape",
                    Geometry = style.Radius > 0 ? "rounded-rectangle" : "rectangle",
                    Radius = style.Radius,
                    Box = box,
                    Fill = style.Fill,
                    LineColor = style.Border,
                    LineWidth = style.BorderWidth,
                    Group = key,
                    Component = style.Identity,
                });
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.bar",
                    Role = "decoration",
                    Type = "shape",
                    Box = LayoutEngine.R(Box.FromSize(box.Left, box.Top + style.Radius, 5, box.Height - (2 * style.Radius))),
                    Fill = sentiment is null ? style.Accent : toneColor,
                    Group = key,
                    Component = style.Identity,
                });

                var value = Scaled(Text(kpi.Value, "kpi_value", "primary", bold: true, heading: true, size: valueSize), factor);
                var valueHeight = MathF.Ceiling(Estimate([value], content.Width) + 2);
                var (valueBox, rest) = LayoutEngine.TakeTop(content, valueHeight, 2);
                Add($"{key}.value", "kpi-value", valueBox, [value]);

                var label = Scaled(Text(kpi.Label, "kpi_label", "muted", size: labelSize), factor);
                var labelHeight = MathF.Ceiling(Estimate([label], content.Width) + 2);
                var (labelBox, afterLabel) = LayoutEngine.TakeTop(rest, labelHeight, 6);
                Add($"{key}.label", "kpi-label", labelBox, [label]);

                if (kpi.Delta is not null)
                {
                    var delta = Scaled(Text(kpi.Delta, "body", toneColor, bold: true, size: deltaSize), factor);
                    var deltaHeight = MathF.Ceiling(Estimate([delta], afterLabel.Width - 16) + 2);
                    var (deltaRow, afterDelta) = LayoutEngine.TakeTop(afterLabel, deltaHeight, 4);
                    if (kpi.Trend is { } trend)
                    {
                        builder.Add(new PlannedElement
                        {
                            Key = $"{key}.arrow",
                            Role = "trend-arrow",
                            Type = "shape",
                            Geometry = "triangle",
                            Box = LayoutEngine.R(Box.FromSize(deltaRow.Left, deltaRow.Top + ((delta.Size * 1.2f) - 9f) / 2f + 1, 10, 9)),
                            Rotation = trend switch { "down" => 180f, "flat" => 90f, _ => 0f },
                            Fill = toneColor,
                            Group = key,
                            Component = style.Identity,
                        });
                    }
                    var deltaBox = kpi.Trend is null ? deltaRow : LayoutEngine.R(new Box(deltaRow.Left + 16, deltaRow.Top, deltaRow.Right, deltaRow.Bottom));
                    Add($"{key}.delta", "kpi-delta", deltaBox, [delta]);
                    afterLabel = afterDelta;
                }
                if (kpi.Note is not null)
                {
                    var note = Scaled(Text(kpi.Note, "caption", "muted"), factor);
                    Add($"{key}.note", "kpi-note", afterLabel, [note]);
                }

                void Add(string elementKey, string role, Box elementBox, IReadOnlyList<PlannedParagraph> paragraphs)
                {
                    builder.Add(new PlannedElement
                    {
                        Key = elementKey,
                        Role = role,
                        Type = "text",
                        Box = elementBox,
                        Paragraphs = paragraphs,
                        Group = key,
                        Component = style.Identity,
                        Region = "kpis",
                    });
                    keys.Add(elementKey);
                }
            }
            builder.Region("kpis", body, "each", 0, keys, splittable: false, Math.Min(MinFont, Profile.Size("caption")));
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public List<PlannedSlide> Appendix()
        {
            var paragraphs = new List<PlannedParagraph>();
            var size = Math.Max(MinFont, Profile.Size("body") - 2f);
            foreach (var reference in Spec.References ?? [])
            {
                var text = reference.Label is null ? reference.Text : $"{reference.Label} {reference.Text}";
                if (reference.Url is not null)
                    text = $"{text} {reference.Url}";
                paragraphs.Add(Text(text, "body", size: size, bullet: reference.Label is null ? "number" : "none", spaceAfter: Profile.ParagraphSpacing, hyperlink: reference.Url));
            }
            foreach (var point in Spec.Points ?? [])
            {
                var text = point.Detail is null ? point.Text : $"{point.Text}: {point.Detail}";
                paragraphs.Add(Text(text, "body", size: size, bullet: "bullet", level: point.Level, spaceAfter: Profile.ParagraphSpacing));
            }

            var probe = Frame();
            float ItemHeight(int index, float factor) => Estimate([Scaled(paragraphs[index], factor)], probe.Body.Width - 24, Profile.LineSpacing) + (paragraphs[index].SpaceAfter * factor);
            var (partition, factor) = PartitionItems(paragraphs.Count, ItemHeight, 0, size, hasTakeaway: false);

            var slides = new List<PlannedSlide>();
            int offset = 0;
            for (int sequence = 0; sequence < partition.Count; sequence++)
            {
                var builder = new SlideBuilder();
                var frame = Frame();
                var title = sequence == 0 ? Spec.Title! : ContinuationTitle(Spec.Title!);
                AddTitle(builder, frame, title);
                var body = sequence == 0 ? AddSubtitle(builder, frame.Body, Spec.Subtitle) : frame.Body;
                builder.Add(new PlannedElement
                {
                    Key = "references",
                    Role = "references",
                    Type = "text",
                    Box = body,
                    Paragraphs = paragraphs.Skip(offset).Take(partition[sequence]).Select(paragraph => Scaled(paragraph, factor)).ToList(),
                    LineSpacing = Profile.LineSpacing,
                    Region = "references",
                });
                builder.Region("references", body, "paragraphs", 0, ["references"], splittable: CanSplit, MinFont);
                AddSource(builder, frame);
                slides.Add(Slide(sequence, builder, title, sequence > 0, offset, partition[sequence]));
                offset += partition[sequence];
            }
            return slides;
        }
    }
}
