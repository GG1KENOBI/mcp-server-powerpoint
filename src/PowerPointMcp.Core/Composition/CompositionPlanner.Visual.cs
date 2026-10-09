using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.Core.Composition;

public static partial class CompositionPlanner
{
    private sealed partial class Planner
    {
        public PlannedSlide TitleSlide()
        {
            var builder = new SlideBuilder();
            var frame = LayoutEngine.Frame(Profile, width, height, hasTitle: false);
            var safe = frame.Safe;
            var titleSize = Profile.Size("title") * 1.35f;
            var titleParagraph = Text(Spec.Title!, "title", heading: true, size: titleSize);
            var titleHeight = MathF.Ceiling(Estimate([titleParagraph], safe.Width * 0.8f) + 8);
            var titleTop = LayoutEngine.Round(safe.Top + (safe.Height * 0.30f));
            var titleBox = LayoutEngine.R(Box.FromSize(safe.Left + 18, titleTop, safe.Width * 0.8f, titleHeight));

            builder.Add(new PlannedElement
            {
                Key = "accent",
                Role = "decoration",
                Type = "shape",
                Box = LayoutEngine.R(Box.FromSize(safe.Left, titleTop + 4, 6, titleHeight - 8)),
                Fill = Profile.Color("primary"),
            });
            builder.Add(TitleElement(frame, Spec.Title!, box: titleBox));
            builder.Elements[^1].Paragraphs = [titleParagraph];
            var keys = new List<string> { "title" };
            var y = titleBox.Bottom + Profile.Spacing("md");
            if (Spec.Subtitle is not null)
            {
                var subtitle = Text(Spec.Subtitle, "subtitle", "muted");
                var subtitleHeight = MathF.Ceiling(Estimate([subtitle], titleBox.Width) + 6);
                var box = LayoutEngine.R(Box.FromSize(titleBox.Left, y, titleBox.Width, subtitleHeight));
                builder.Add(new PlannedElement { Key = "subtitle", Role = "subtitle", Type = "text", Box = box, Paragraphs = [subtitle], Region = "title-block" });
                keys.Add("subtitle");
            }
            var meta = new List<PlannedParagraph>();
            if (Spec.Author is not null)
                meta.Add(Text(Spec.Author, "body", bold: true, heading: true));
            if (Spec.Date is not null)
                meta.Add(Text(Spec.Date, "body", "muted"));
            if (meta.Count > 0)
            {
                var metaHeight = MathF.Ceiling(Estimate(meta, safe.Width / 2f) + 8);
                var box = LayoutEngine.R(new Box(titleBox.Left, safe.Bottom - metaHeight, titleBox.Left + (safe.Width / 2f), safe.Bottom));
                builder.Add(new PlannedElement { Key = "meta", Role = "byline", Type = "text", Box = box, Paragraphs = meta, VAlign = "bottom", Region = "title-block" });
                keys.Add("meta");
            }
            builder.Region("title-block", safe, "each", 0, keys, splittable: false, MinFont);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide SectionSlide()
        {
            var builder = new SlideBuilder();
            var frame = LayoutEngine.Frame(Profile, width, height, hasTitle: false);
            var style = Profile.Component("section-header");
            builder.Add(new PlannedElement
            {
                Key = "background",
                Role = "background",
                Type = "shape",
                Box = new Box(0, 0, width, height),
                Fill = style.Fill,
                Component = style.Identity,
            });
            var safe = frame.Safe;
            var top = LayoutEngine.Round(safe.Top + (safe.Height * 0.36f));
            var left = safe.Left + style.Padding;
            var textWidth = safe.Width * 0.75f;
            if (Spec.SectionNumber is not null)
            {
                var number = Text(Spec.SectionNumber, "title", style.Accent, bold: true, heading: true, size: style.HeadingSize * 1.4f);
                var numberHeight = MathF.Ceiling(Estimate([number], textWidth) + 4);
                builder.Add(new PlannedElement
                {
                    Key = "number",
                    Role = "section-number",
                    Type = "text",
                    Box = LayoutEngine.R(Box.FromSize(left, top - numberHeight - 4, textWidth, numberHeight)),
                    Paragraphs = [number],
                    VAlign = "bottom",
                    Component = style.Identity,
                });
            }
            var title = Text(Spec.Title!, "title", style.Text, heading: true, size: style.HeadingSize * 1.2f);
            var titleHeight = MathF.Ceiling(Estimate([title], textWidth) + 8);
            var titleBox = LayoutEngine.R(Box.FromSize(left, top, textWidth, titleHeight));
            builder.Add(TitleElement(frame, Spec.Title!, box: titleBox, vAlign: "top"));
            builder.Elements[^1].Paragraphs = [title];
            var keys = new List<string> { "title" };
            if (Spec.Subtitle is not null)
            {
                var subtitle = Text(Spec.Subtitle, "subtitle", style.Text, size: style.BodySize);
                var subtitleHeight = MathF.Ceiling(Estimate([subtitle], textWidth) + 6);
                builder.Add(new PlannedElement
                {
                    Key = "subtitle",
                    Role = "subtitle",
                    Type = "text",
                    Box = LayoutEngine.R(Box.FromSize(left, titleBox.Bottom + Profile.Spacing("sm"), textWidth, subtitleHeight)),
                    Paragraphs = [subtitle],
                    Region = "section",
                    Component = style.Identity,
                });
                keys.Add("subtitle");
            }
            builder.Region("section", safe, "each", 0, keys, splittable: false, MinFont);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide ImageText()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var image = Spec.Image!;
            var columns = LayoutEngine.Columns(frame.Body, 2, Profile.Gutter * 1.5f, [0.48f, 0.52f]);
            var (imageColumn, textColumn) = image.Position == "right" ? (columns[1], columns[0]) : (columns[0], columns[1]);
            if (image.Position == "right")
            {
                // Keep the text column on the left at 52% width.
                var swapped = LayoutEngine.Columns(frame.Body, 2, Profile.Gutter * 1.5f, [0.52f, 0.48f]);
                (textColumn, imageColumn) = (swapped[0], swapped[1]);
            }

            var pictureArea = imageColumn;
            if (image.Attribution is not null)
            {
                var (captionBox, above) = LayoutEngine.TakeBottom(imageColumn, Profile.Size("caption") * 1.4f, 4);
                pictureArea = above;
                builder.Add(new PlannedElement
                {
                    Key = "image.attribution",
                    Role = "attribution",
                    Type = "text",
                    Box = captionBox,
                    Paragraphs = [Text(image.Attribution, "caption", "muted", size: Math.Max(Profile.Size("caption"), Math.Min(MinFont, Profile.Size("caption"))))],
                    Region = "attribution",
                });
                builder.Region("attribution", captionBox, "each", 0, ["image.attribution"], splittable: false, Math.Min(MinFont, Profile.Size("footnote")));
            }

            if (!Context.ImageSizes.TryGetValue(image.Path, out var pixels))
                throw new ArgumentException($"Image size for '{image.Path}' was not provided to the planner.");
            var placement = PictureFit.Compute(pictureArea, pixels.Width / (double)pixels.Height, image.Fit, image.FocalX, image.FocalY);
            builder.Add(new PlannedElement
            {
                Key = "image",
                Role = "picture",
                Type = "picture",
                Box = placement.Frame,
                Picture = new PicturePlan(image.Path, placement, image.Alt, pixels.Width, pixels.Height, image.Fit, image.Attribution),
                AltText = image.Alt,
            });

            var text = AddSubtitle(builder, textColumn, Spec.Subtitle);
            text = AddTakeaway(builder, text, Spec.Takeaway);
            if (Spec.Points is { Count: > 0 } points)
            {
                IReadOnlyList<PlannedParagraph> Bullets(float factor) => points.Select(point => Scaled(Text(
                    point.Detail is null ? point.Text : $"{point.Text}: {point.Detail}", "body", bullet: "bullet", level: point.Level,
                    spaceAfter: Profile.ParagraphSpacing), factor)).ToList();
                var factor = LargestFittingFactor(Profile.Size("body"), f => Estimate(Bullets(f), text.Width - 20, Profile.LineSpacing) <= text.Height);
                if (factor is null)
                    Warnings.Add("The points beside the image are estimated to overflow; shorten them or use a table or appendix.");
                builder.Add(new PlannedElement
                {
                    Key = "points",
                    Role = "body",
                    Type = "text",
                    Box = text,
                    Paragraphs = Bullets(factor ?? Math.Min(1f, MinFont / Profile.Size("body"))),
                    LineSpacing = Profile.LineSpacing,
                    Region = "text",
                });
                builder.Region("text", text, "each", 0, ["points"], splittable: false, MinFont);
            }
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public List<PlannedSlide> Table()
        {
            var table = Spec.Table!;
            var style = Profile.Table;
            var probe = Frame();
            var fontSize = style.FontSize;
            var align = table.Align ?? Enumerable.Range(0, table.Header.Count)
                .Select(column => table.Rows.Where(row => !string.IsNullOrWhiteSpace(row[column])).All(row => LooksNumeric(row[column])) &&
                    table.Rows.Any(row => !string.IsNullOrWhiteSpace(row[column])) ? "right" : "left")
                .ToList();

            var widths = ColumnWidths(table, probe.Body.Width, fontSize);
            float RowHeight(IReadOnlyList<string> cells, float size, bool bold) =>
                MathF.Ceiling(cells.Select((cell, column) => TextEstimator.Height(cell, size, widths[column] - (2 * style.CellPadding) - 2, TextEstimator.DefaultLineHeight, 0, bold)).DefaultIfEmpty(size).Max() + (2 * style.CellPadding) + 4);

            var headerHeight = RowHeight(table.Header, fontSize, bold: true);
            float ItemHeight(int index, float factor) => RowHeight(table.Rows[index], fontSize * factor, table.TotalRow && index == table.Rows.Count - 1);
            float HeaderAt(float factor) => RowHeight(table.Header, fontSize * factor, bold: true);

            // Treat the header as part of every slide's fixed cost by reserving it from the item area.
            var smallest = fontSize;
            var originalFrame = probe;
            var takeaway = Spec.Takeaway is null ? 0f : TakeawayMinHeight + Profile.Spacing("md");
            var subtitle = Spec.Subtitle is null ? 0f : Estimate([Text(Spec.Subtitle, "subtitle")], originalFrame.Body.Width) + 4 + Profile.Spacing("sm");
            List<int> partition;
            float factor = 1f;
            if (Context.Partition is { } forced)
            {
                partition = [.. forced];
            }
            else
            {
                float Total(int start, int length, float f) => HeaderAt(f) + Enumerable.Range(start, length).Sum(index => ItemHeight(index, f));
                var single = originalFrame.Body.Height - takeaway - subtitle;
                var fitting = LargestFittingFactor(smallest, f => Total(0, table.Rows.Count, f) <= single);
                if (fitting is { } value)
                {
                    partition = [table.Rows.Count];
                    factor = value;
                }
                else if (CanSplit && table.Rows.Count > 1)
                {
                    partition = [];
                    int start = 0;
                    while (start < table.Rows.Count)
                    {
                        bool first = partition.Count == 0;
                        int take = 0;
                        while (start + take < table.Rows.Count &&
                            Total(start, take + 1, 1f) <= originalFrame.Body.Height - (first ? subtitle : 0) - (start + take + 1 == table.Rows.Count ? takeaway : 0))
                        {
                            take++;
                        }
                        take = Math.Max(1, take);
                        partition.Add(take);
                        start += take;
                    }
                    Warnings.Add($"The table is estimated not to fit one slide; planned {partition.Count} slides with the header repeated ({string.Join(" + ", partition)} rows).");
                }
                else
                {
                    partition = [table.Rows.Count];
                    factor = CanShrink ? Math.Min(1f, MinFont / fontSize) : 1f;
                    Warnings.Add("The table is estimated to overflow and splitting is not allowed by fit.policy; overflow will be reported.");
                }
            }

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
                var rows = table.Rows.Skip(offset).Take(partition[sequence]).ToList();
                var isLast = sequence == partition.Count - 1;
                var highlight = (table.Highlight ?? []).Where(row => row > offset && row <= offset + rows.Count).Select(row => row - offset).ToList();
                var estimated = HeaderAt(factor) + rows.Select((_, i) => ItemHeight(offset + i, factor)).Sum();
                var tableBox = LayoutEngine.R(new Box(body.Left, body.Top, body.Right, Math.Min(body.Bottom, body.Top + estimated)));
                var header = sequence == 0 || table.RepeatHeader ? table.Header : null;
                builder.Add(new PlannedElement
                {
                    Key = "table",
                    Role = "table",
                    Type = "table",
                    Box = tableBox,
                    Table = new TablePlan(header ?? [], rows, widths, align, table.TotalRow && isLast, highlight, LayoutEngine.Round(fontSize * factor), Profile.BodyFont, style, offset + 1),
                    Region = "table",
                    EstimatedTextHeight = estimated,
                });
                builder.Region("table", body, "table", 0, ["table"], splittable: CanSplit, MinFont);
                AddSource(builder, frame);
                slides.Add(Slide(sequence, builder, title, sequence > 0, offset, partition[sequence]));
                offset += partition[sequence];
            }
            return slides;
        }

        private static List<float> ColumnWidths(SpecTable table, float totalWidth, float fontSize)
        {
            if (table.ColumnWeights is { } weights)
            {
                var sum = weights.Sum();
                return weights.Select(weight => LayoutEngine.Round(totalWidth * weight / sum)).ToList();
            }
            // Natural width of the longest line per column, softened so long text columns wrap.
            var natural = Enumerable.Range(0, table.Header.Count).Select(column =>
            {
                var cells = table.Rows.Select(row => row[column]).Append(table.Header[column]);
                var longest = cells.Max(cell => TextEstimator.LineWidth(cell, fontSize, bold: false)) + 16f;
                return MathF.Sqrt(Math.Max(40f, longest)) * MathF.Sqrt(Math.Min(longest, 400f));
            }).ToList();
            var totalNatural = natural.Sum();
            return natural.Select(value => LayoutEngine.Round(totalWidth * value / totalNatural)).ToList();
        }

        public PlannedSlide Chart()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            var chartArea = body;
            if (Spec.Takeaway is not null)
            {
                var columns = LayoutEngine.Columns(body, 2, Profile.Gutter * 1.5f, [0.68f, 0.32f]);
                chartArea = columns[0];
                AddTakeaway(builder, columns[1], Spec.Takeaway, vertical: true);
            }
            var chart = Spec.Chart!;
            builder.Add(new PlannedElement
            {
                Key = "chart",
                Role = "chart",
                Type = "chart",
                Box = chartArea,
                Chart = new ChartPlan(chart, Profile.Chart.Palette, Profile.BodyFont, Profile.Chart.FontSize, Profile.Color("text"),
                    Profile.Chart.Gridlines, chart.DataLabels ?? Profile.Chart.DataLabels,
                    chart.Legend == false ? "none" : chart.Series.Count == 1 && chart.Type is not ("pie" or "doughnut") && chart.Legend != true ? "none" : Profile.Chart.Legend),
                AltText = $"{chart.Type} chart: {Spec.Title}",
            });
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide Timeline()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            body = AddTakeaway(builder, body, Spec.Takeaway);
            var milestones = Spec.Milestones!;
            var style = Profile.Component("timeline");
            int count = milestones.Count;
            bool alternate = count > 6;
            var slot = body.Width / count;
            var axisY = LayoutEngine.Round(body.Top + (body.Height * (alternate ? 0.5f : 0.34f)));
            var labelWidth = LayoutEngine.Round(Math.Min(alternate ? (2 * slot) - Profile.Gutter : slot - Profile.Gutter, 220f));

            builder.Add(new PlannedElement
            {
                Key = "axis",
                Role = "timeline-axis",
                Type = "line",
                Box = LayoutEngine.R(new Box(body.Left, axisY, body.Right, axisY)),
                LineColor = style.Border ?? Profile.Color("border"),
                LineWidth = style.BorderWidth,
                Component = style.Identity,
            });

            var keys = new List<string>();
            for (int i = 0; i < count; i++)
            {
                var milestone = milestones[i];
                var key = $"milestone-{N(i + 1)}";
                var centerX = body.Left + (slot * (i + 0.5f));
                var markerSize = milestone.Status == "current" ? 22f : 16f;
                var (fill, line) = milestone.Status switch
                {
                    "done" => (style.Accent, style.Accent),
                    "current" => (Profile.Color("accent"), Profile.Color("accent")),
                    _ => (Profile.Color("background"), style.Accent),
                };
                builder.Add(new PlannedElement
                {
                    Key = $"{key}.marker",
                    Role = "timeline-marker",
                    Type = "shape",
                    Geometry = "oval",
                    Box = LayoutEngine.R(Box.FromSize(centerX - (markerSize / 2f), axisY - (markerSize / 2f), markerSize, markerSize)),
                    Fill = fill,
                    LineColor = line,
                    LineWidth = 2,
                    Component = style.Identity,
                    Tags = new Dictionary<string, string>(StringComparer.Ordinal) { ["PPTMCP_STATUS"] = milestone.Status },
                });

                bool above = alternate && i % 2 == 1;
                var date = Text(milestone.Date, "body", "primary", bold: true, heading: true, align: "center", size: style.HeadingSize);
                var label = new List<PlannedParagraph> { Text(milestone.Label, "body", bold: true, align: "center", size: style.HeadingSize) };
                if (milestone.Detail is not null)
                    label.Add(Text(milestone.Detail, "caption", "muted", align: "center", size: style.BodySize));
                var dateHeight = MathF.Ceiling(Estimate([date], labelWidth) + 4);
                var labelHeight = alternate
                    ? (body.Height / 2f) - dateHeight - 26f
                    : body.Bottom - axisY - 18f - dateHeight - 6f;
                Box dateBox, labelBox;
                if (!alternate)
                {
                    dateBox = LayoutEngine.R(Box.FromSize(centerX - (labelWidth / 2f), axisY - 16f - dateHeight, labelWidth, dateHeight));
                    labelBox = LayoutEngine.R(Box.FromSize(centerX - (labelWidth / 2f), axisY + 18f, labelWidth, Math.Max(24f, body.Bottom - axisY - 18f)));
                }
                else if (!above)
                {
                    dateBox = LayoutEngine.R(Box.FromSize(centerX - (labelWidth / 2f), axisY + 16f, labelWidth, dateHeight));
                    labelBox = LayoutEngine.R(Box.FromSize(centerX - (labelWidth / 2f), dateBox.Bottom + 4f, labelWidth, Math.Max(24f, labelHeight)));
                }
                else
                {
                    dateBox = LayoutEngine.R(Box.FromSize(centerX - (labelWidth / 2f), axisY - 16f - dateHeight, labelWidth, dateHeight));
                    labelBox = LayoutEngine.R(new Box(centerX - (labelWidth / 2f), Math.Max(body.Top, dateBox.Top - 4f - Math.Max(24f, labelHeight)), centerX + (labelWidth / 2f), dateBox.Top - 4f));
                }
                builder.Add(new PlannedElement { Key = $"{key}.date", Role = "timeline-date", Type = "text", Box = dateBox, Paragraphs = [date], VAlign = above || !alternate ? "bottom" : "top", Component = style.Identity, Region = "milestones" });
                builder.Add(new PlannedElement { Key = $"{key}.label", Role = "timeline-label", Type = "text", Box = labelBox, Paragraphs = label, VAlign = above ? "bottom" : "top", Component = style.Identity, Region = "milestones" });
                keys.Add($"{key}.date");
                keys.Add($"{key}.label");
            }
            builder.Region("milestones", body, "each", 0, keys, splittable: false, Math.Min(MinFont, style.BodySize));
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide Process()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            body = AddTakeaway(builder, body, Spec.Takeaway);
            var steps = Spec.Steps!;
            var style = Profile.Component("process-step");
            var columns = LayoutEngine.Columns(body, steps.Count, 6f);
            var chevronHeight = MathF.Ceiling(Math.Max(56f, (style.HeadingSize * 2.6f) + (2 * style.Padding)));
            var keys = new List<string>();
            var detailKeys = new List<string>();
            for (int i = 0; i < steps.Count; i++)
            {
                var step = steps[i];
                var key = $"step-{N(i + 1)}";
                var column = columns[i];
                var notch = Math.Min(chevronHeight * 0.35f, column.Width * 0.2f);
                var shape = LayoutEngine.R(Box.FromSize(column.Left, body.Top, column.Width, chevronHeight));
                builder.Add(new PlannedElement
                {
                    Key = key,
                    Role = "process-step",
                    Type = "shape",
                    Geometry = i == 0 ? "pentagon" : "chevron",
                    Box = shape,
                    Fill = i % 2 == 0 ? style.Fill : style.Accent,
                    Paragraphs = [Text(step.Label, "body", style.Text, bold: true, heading: true, align: "center", size: style.HeadingSize)],
                    VAlign = "middle",
                    Padding = [i == 0 ? style.Padding : notch + 4, 4, notch + 4, 4],
                    Component = style.Identity,
                    Region = "steps",
                });
                keys.Add(key);
                if (step.Detail is not null)
                {
                    var detailBox = LayoutEngine.R(new Box(column.Left + 4, shape.Bottom + Profile.Spacing("md"), column.Right - 4, body.Bottom));
                    builder.Add(new PlannedElement
                    {
                        Key = $"{key}.detail",
                        Role = "process-detail",
                        Type = "text",
                        Box = detailBox,
                        Paragraphs = [Text(step.Detail, "body", "muted", align: "center", size: style.BodySize)],
                        Component = style.Identity,
                        Region = "details",
                    });
                    detailKeys.Add($"{key}.detail");
                }
            }
            builder.Region("steps", body, "each", 0, keys, splittable: false, Math.Min(MinFont, style.BodySize));
            if (detailKeys.Count > 0)
                builder.Region("details", body, "each", 0, detailKeys, splittable: false, Math.Min(MinFont, style.BodySize));
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        public PlannedSlide Hierarchy()
        {
            var builder = new SlideBuilder();
            var frame = Frame();
            AddTitle(builder, frame, Spec.Title);
            var body = AddSubtitle(builder, frame.Body, Spec.Subtitle);
            var style = Profile.Component("node");
            var hasDetail = HasDetail(Spec.Tree!);
            var nodeHeight = MathF.Ceiling((style.HeadingSize * 1.3f) + (hasDetail ? style.BodySize * 1.3f : 0) + (2 * style.Padding));
            var layout = TreeLayout.Layout(Spec.Tree!, body, nodeHeight, 64f, 190f, Profile.Spacing("sm"));
            float nodeFontFactor = 1f;
            if (layout is null && CanShrink)
            {
                nodeFontFactor = Math.Max(MinFont / style.HeadingSize, 0.75f);
                nodeHeight = MathF.Ceiling(nodeHeight * nodeFontFactor);
                layout = TreeLayout.Layout(Spec.Tree!, body, nodeHeight, 48f, 190f, Profile.Spacing("xs"));
            }
            if (layout is null)
                throw new ArgumentException($"The hierarchy has {TreeLayout.CountLeaves(Spec.Tree!)} leaves and {TreeLayout.Depth(Spec.Tree!)} levels, which do not fit on one slide. Split it into sub-trees (one slide per branch).");

            var keys = new List<string>();
            foreach (var node in layout)
            {
                var key = $"node-{node.Id}";
                var paragraphs = new List<PlannedParagraph> { Scaled(Text(node.Label, "body", style.Text, bold: true, heading: true, align: "center", size: style.HeadingSize), nodeFontFactor) };
                if (node.Detail is not null)
                    paragraphs.Add(Scaled(Text(node.Detail, "caption", "muted", align: "center", size: style.BodySize), nodeFontFactor));
                builder.Add(new PlannedElement
                {
                    Key = key,
                    Role = node.Depth == 0 ? "node-root" : "node",
                    Type = "shape",
                    Geometry = style.Radius > 0 ? "rounded-rectangle" : "rectangle",
                    Radius = style.Radius,
                    Box = node.Box,
                    Fill = node.Depth == 0 ? Profile.Color("primary") : style.Fill,
                    LineColor = style.Border,
                    LineWidth = style.BorderWidth,
                    Paragraphs = node.Depth == 0
                        ? paragraphs.Select(paragraph => new PlannedParagraph { Text = paragraph.Text, Size = paragraph.Size, Font = paragraph.Font, Color = Profile.Color("background"), Bold = paragraph.Bold, Align = "center" }).ToList()
                        : paragraphs,
                    VAlign = "middle",
                    Padding = [4, 2, 4, 2],
                    Component = style.Identity,
                    Region = "nodes",
                    Tags = new Dictionary<string, string>(StringComparer.Ordinal) { [DeckRoles.NodeTag] = node.Id },
                });
                keys.Add(key);
            }
            foreach (var node in layout.Where(node => node.ParentId is not null))
            {
                builder.Add(new PlannedElement
                {
                    Key = $"edge-{node.ParentId}-{node.Id}",
                    Role = "connector",
                    Type = "connector",
                    Box = node.Box,
                    FromKey = $"node-{node.ParentId}",
                    ToKey = $"node-{node.Id}",
                    FromSite = 3,
                    ToSite = 1,
                    ConnectorKind = "elbow",
                    LineColor = style.Border ?? Profile.Color("border"),
                    LineWidth = 1.25f,
                    Component = style.Identity,
                });
            }
            builder.Region("nodes", body, "each", 0, keys, splittable: false, MinFont * 0.85f);
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }

        private static bool HasDetail(SpecNode node) => node.Detail is not null || node.Children.Any(HasDetail);

        public PlannedSlide Quote()
        {
            var builder = new SlideBuilder();
            var hasTitle = Spec.Title is not null;
            var frame = Frame(hasTitle);
            AddTitle(builder, frame, Spec.Title);
            var quote = Spec.Quote!;
            var style = Profile.Component("quote");
            var body = frame.Body;
            var markSize = style.HeadingSize * 3f;
            builder.Add(new PlannedElement
            {
                Key = "quote.mark",
                Role = "decoration",
                Type = "text",
                Box = LayoutEngine.R(Box.FromSize(body.Left, body.Top, markSize * 0.55f, markSize * 1.1f)),
                Paragraphs = [Text("“", "quote", style.Accent, bold: true, heading: true, size: markSize)],
                Component = style.Identity,
            });
            var textLeft = body.Left + (markSize * 0.55f) + Profile.Spacing("sm");
            var textBox = new Box(textLeft, body.Top + (markSize * 0.35f), body.Right - (body.Width * 0.08f), body.Bottom);
            var quoteParagraph = Text(quote.Text, "quote", style.Text, italic: true, heading: true, size: style.HeadingSize, spaceAfter: Profile.Spacing("md"));
            var attribution = new List<PlannedParagraph> { Text(quote.Author, "body", bold: true, heading: true) };
            if (quote.Role is not null)
                attribution.Add(Text(quote.Role, "caption", "muted"));
            var attributionHeight = MathF.Ceiling(Estimate(attribution, textBox.Width) + 6);
            var (authorBox, quoteBox) = LayoutEngine.TakeBottom(textBox, attributionHeight, Profile.Spacing("md"));
            var factor = LargestFittingFactor(style.HeadingSize, f => Estimate([Scaled(quoteParagraph, f)], quoteBox.Width, 1.1f) <= quoteBox.Height) ?? 1f;
            builder.Add(new PlannedElement
            {
                Key = "quote.text",
                Role = "quote",
                Type = "text",
                Box = quoteBox,
                Paragraphs = [Scaled(quoteParagraph, factor)],
                LineSpacing = 1.1f,
                Component = style.Identity,
                Region = "quote",
            });
            builder.Add(new PlannedElement
            {
                Key = "quote.author",
                Role = "attribution",
                Type = "text",
                Box = authorBox,
                Paragraphs = attribution,
                Component = style.Identity,
                Region = "quote",
            });
            builder.Region("quote", body, "each", 0, ["quote.text", "quote.author"], splittable: false, MinFont);
            AddSource(builder, frame);
            return Slide(0, builder, Spec.Title);
        }
    }
}
