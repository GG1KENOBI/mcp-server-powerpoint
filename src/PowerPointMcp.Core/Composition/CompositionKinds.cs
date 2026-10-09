namespace Sbroenne.PowerPointMcp.Core.Composition;

/// <summary>One composition kind: what it builds, which fields it uses, and an example.</summary>
public sealed class CompositionKind
{
    /// <summary>Kind name.</summary>
    public required string Name { get; init; }

    /// <summary>What the slide looks like.</summary>
    public required string Description { get; init; }

    /// <summary>Required fields.</summary>
    public required IReadOnlyList<string> Required { get; init; }

    /// <summary>Optional fields.</summary>
    public required IReadOnlyList<string> Optional { get; init; }

    /// <summary>A complete example.</summary>
    public required string Example { get; init; }

    internal Action<CompositionSpec, List<string>, List<string>> Check { get; init; } = (_, _, _) => { };

    internal void Validate(CompositionSpec spec, List<string> errors, List<string> warnings)
    {
        foreach (var field in Required.Where(field => !Has(spec, field)))
            errors.Add($"$.{field}: required for kind {Name}.");
        foreach (var field in Present(spec).Where(field => !Required.Contains(field) && !Optional.Contains(field) && !Common.Contains(field)))
            warnings.Add($"$.{field}: not used by kind {Name}; ignored.");
        Check(spec, errors, warnings);
    }

    private static readonly string[] Common = ["kind", "id", "insert_at", "profile", "notes", "fit", "metadata"];

    private static bool Has(CompositionSpec spec, string field) => Present(spec).Contains(field);

    private static List<string> Present(CompositionSpec spec)
    {
        var fields = new List<string>();
        void Add(string name, object? value)
        {
            if (value is not null)
                fields.Add(name);
        }
        Add("title", spec.Title);
        Add("subtitle", spec.Subtitle);
        Add("source", spec.Source);
        Add("takeaway", spec.Takeaway);
        Add("date", spec.Date);
        Add("author", spec.Author);
        Add("section_number", spec.SectionNumber);
        Add("points", spec.Points);
        Add("columns", spec.Columns);
        Add("cards", spec.Cards);
        Add("kpis", spec.Kpis);
        Add("image", spec.Image);
        Add("table", spec.Table);
        Add("chart", spec.Chart);
        Add("milestones", spec.Milestones);
        Add("steps", spec.Steps);
        Add("tree", spec.Tree);
        Add("quote", spec.Quote);
        Add("references", spec.References);
        return fields;
    }
}

/// <summary>The composition kinds.</summary>
public static class CompositionKinds
{
    /// <summary>Chart types compositions create.</summary>
    public static readonly string[] ChartTypes = ["column", "bar", "line", "pie", "doughnut", "area", "stacked-column", "stacked-bar"];

    /// <summary>All kinds.</summary>
    public static readonly IReadOnlyList<CompositionKind> All =
    [
        new()
        {
            Name = "title",
            Description = "Title slide: large title, subtitle, presenter and date lines.",
            Required = ["title"],
            Optional = ["subtitle", "author", "date"],
            Example = """{"kind":"title","title":"Итоги 2025 года","subtitle":"Совет директоров","author":"Финансовый департамент","date":"Март 2026"}""",
        },
        new()
        {
            Name = "section",
            Description = "Section divider: full-bleed color band with section number and name.",
            Required = ["title"],
            Optional = ["section_number", "subtitle"],
            Example = """{"kind":"section","section_number":"02","title":"Market review","subtitle":"Demand, pricing, competitors"}""",
        },
        new()
        {
            Name = "executive-summary",
            Description = "Numbered key points (headline + optional detail) with an optional takeaway banner. Splits onto continuation slides when the points do not fit.",
            Required = ["title", "points"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"executive-summary","title":"Executive summary","points":[{"text":"Revenue up 12%","detail":"Driven by EMEA enterprise deals"},"Margin up 3 points","Cash flow positive since Q2"],"takeaway":"On track for the 2026 plan"}""",
            Check = (spec, errors, _) => Count(spec.Points, "points", 1, 20, errors),
        },
        new()
        {
            Name = "comparison",
            Description = "Two columns with headings and points (e.g. options, before/after, pros/cons) and an optional verdict.",
            Required = ["title", "columns"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"comparison","title":"Build or buy","columns":[{"heading":"Build","points":["Full control","12 months"],"tone":"neutral"},{"heading":"Buy","points":["Live in 8 weeks","Licence cost"],"tone":"positive"}],"takeaway":"Buy, then extend"}""",
            Check = (spec, errors, _) =>
            {
                if (spec.Columns is { Count: not 2 })
                    errors.Add($"$.columns: a comparison has exactly 2 columns (got {spec.Columns.Count}).");
            },
        },
        new()
        {
            Name = "cards",
            Description = "Equal cards in a grid (1-8), each with heading, body, and optional badge.",
            Required = ["title", "cards"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"cards","title":"Our priorities","cards":[{"heading":"Growth","body":"Enter two new markets","badge":"P1"},{"heading":"Efficiency","body":"Automate reporting"},{"heading":"People","body":"Hire 40 engineers"}]}""",
            Check = (spec, errors, _) => Count(spec.Cards, "cards", 1, 8, errors),
        },
        new()
        {
            Name = "kpis",
            Description = "KPI panels (1-8): big value, label, delta with a trend arrow (trend = direction, sentiment = color), and note.",
            Required = ["title", "kpis"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"kpis","title":"Q3 at a glance","kpis":[{"value":"$4.2M","label":"Revenue","delta":"+12% YoY","trend":"up"},{"value":"38%","label":"Gross margin","delta":"+3 pts","trend":"up"},{"value":"2.1%","label":"Churn","delta":"+0.4 pts","trend":"up","sentiment":"negative"}],"source":"Finance, unaudited"}""",
            Check = (spec, errors, _) => Count(spec.Kpis, "kpis", 1, 8, errors),
        },
        new()
        {
            Name = "image-text",
            Description = "A local picture (cover or contain, focal point) beside points; attribution under the picture.",
            Required = ["title", "image"],
            Optional = ["subtitle", "points", "takeaway", "source"],
            Example = """{"kind":"image-text","title":"New plant in Kazan","image":{"path":"C:\\Assets\\plant.jpg","alt":"Aerial view of the plant","fit":"cover","position":"left","attribution":"Photo: company archive"},"points":["Opened in May","1,200 jobs"]}""",
            Check = (spec, errors, _) =>
            {
                if (spec.Points is null && spec.Takeaway is null && spec.Subtitle is null)
                    errors.Add("$.points: image-text needs points, a takeaway, or a subtitle beside the picture.");
            },
        },
        new()
        {
            Name = "table",
            Description = "A native table styled by the profile with an optional takeaway; long tables continue on new slides with the header repeated.",
            Required = ["title", "table"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"table","title":"Revenue by region","table":{"header":["Region","2024","2025","Change"],"rows":[["EMEA","3.1","3.6","+16%"],["APAC","2.4","2.9","+21%"],["Total","5.5","6.5","+18%"]],"align":["left","right","right","right"],"total_row":true},"takeaway":"APAC grows fastest","source":"Management accounts"}""",
        },
        new()
        {
            Name = "chart",
            Description = "A native chart (column, bar, line, pie, doughnut, area, stacked) with profile colors and a takeaway beside it.",
            Required = ["title", "chart"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"chart","title":"Revenue trend","chart":{"type":"column","categories":["Q1","Q2","Q3","Q4"],"series":[{"name":"2024","values":[1.1,1.3,1.2,1.5]},{"name":"2025","values":[1.3,1.5,1.6,null]}],"value_axis_title":"$M"},"takeaway":"Q3 record despite seasonality","source":"Finance"}""",
        },
        new()
        {
            Name = "timeline",
            Description = "Horizontal timeline with milestones (2-12) alternating above and below the axis; status colors done/current/planned.",
            Required = ["title", "milestones"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"timeline","title":"Roadmap 2026","milestones":[{"date":"Jan","label":"Pilot","status":"done"},{"date":"Apr","label":"Rollout wave 1","status":"current"},{"date":"Sep","label":"Wave 2"},{"date":"Dec","label":"Full scale"}]}""",
            Check = (spec, errors, _) => Count(spec.Milestones, "milestones", 2, 12, errors),
        },
        new()
        {
            Name = "process",
            Description = "Left-to-right process (2-8 steps) as connected chevrons with details below.",
            Required = ["title", "steps"],
            Optional = ["subtitle", "takeaway", "source"],
            Example = """{"kind":"process","title":"How onboarding works","steps":[{"label":"Apply","detail":"Online form"},{"label":"Review","detail":"2 days"},{"label":"Sign","detail":"E-signature"},{"label":"Start"}]}""",
            Check = (spec, errors, _) => Count(spec.Steps, "steps", 2, 8, errors),
        },
        new()
        {
            Name = "hierarchy",
            Description = "Org chart / hierarchy tree (up to 4 levels, 40 nodes) with elbow connectors glued to the boxes.",
            Required = ["title", "tree"],
            Optional = ["subtitle", "source"],
            Example = """{"kind":"hierarchy","title":"Project organisation","tree":{"label":"Steering committee","children":[{"label":"PMO","children":[{"label":"Finance"},{"label":"IT"}]},{"label":"Business lead"}]}}""",
            Check = (spec, errors, _) =>
            {
                if (spec.Tree is not null && CountNodes(spec.Tree) > 40)
                    errors.Add("$.tree: more than 40 nodes do not fit on one slide; split the hierarchy.");
            },
        },
        new()
        {
            Name = "quote",
            Description = "A large quotation with author and role.",
            Required = ["quote"],
            Optional = ["title", "source"],
            Example = """{"kind":"quote","quote":{"text":"The best way to predict the future is to create it.","author":"Peter Drucker","role":"Management consultant"}}""",
        },
        new()
        {
            Name = "appendix",
            Description = "Appendix or references: numbered references (with URLs as hyperlinks) and/or points in compact type, continuing onto new slides as needed.",
            Required = ["title"],
            Optional = ["references", "points", "subtitle", "source"],
            Example = """{"kind":"appendix","title":"Sources","references":[{"label":"[1]","text":"Rosstat, Socio-economic situation of Russia, 2025","url":"https://rosstat.gov.ru"},"Company management accounts, FY2025"]}""",
            Check = (spec, errors, _) =>
            {
                if (spec.References is null && spec.Points is null)
                    errors.Add("$.references: appendix needs references or points.");
            },
        },
    ];

    /// <summary>Finds a kind by name.</summary>
    public static CompositionKind? Find(string name) => All.FirstOrDefault(kind => kind.Name == name);

    private static void Count<T>(IReadOnlyList<T>? items, string name, int min, int max, List<string> errors)
    {
        if (items is null)
            return;
        if (items.Count < min || items.Count > max)
            errors.Add($"$.{name}: needs {min}-{max} items (got {items.Count}).");
    }

    private static int CountNodes(SpecNode node) => 1 + node.Children.Sum(CountNodes);
}
