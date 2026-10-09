// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using Sbroenne.PowerPointMcp.Core.Deck;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Deck;

/// <summary>Builders for inspection snapshots used by the pure deck tests.</summary>
internal static class DeckTestData
{
    internal static DeckSlideInfo Slide(int index, int id, string? title = null, string? appId = null, string? section = null,
        string? layout = null, IReadOnlyDictionary<string, string>? tags = null) => new()
        {
            SlideIndex = index,
            SlideId = id,
            Title = title,
            AppId = appId,
            SectionName = section,
            LayoutName = layout,
            Hidden = false,
            ShapeCount = 0,
            Tags = tags,
        };

    internal static DeckObjectInfo Shape(int slideIndex, int slideId, int shapeId, string name, string kind = "text-box",
        string? text = null, string? role = null, string? appId = null, float left = 10, float top = 10, float width = 100,
        float height = 50, float? minFont = null, int? group = null, bool visible = true, string? placeholder = null,
        string? component = null, IReadOnlyDictionary<string, string>? tags = null) => new()
        {
            SlideIndex = slideIndex,
            SlideId = slideId,
            ShapeIndex = shapeId,
            ShapeId = shapeId,
            Name = name,
            Kind = kind,
            Role = role,
            AppId = appId,
            Left = left,
            Top = top,
            Width = width,
            Height = height,
            Rotation = 0,
            ZOrder = shapeId,
            Visible = visible,
            HasText = text is not null,
            Text = text,
            MinFontSize = minFont,
            MaxFontSize = minFont,
            GroupShapeId = group,
            PlaceholderType = placeholder,
            Component = component,
            Tags = tags,
        };

    /// <summary>Two slides: an agenda with a title and two cards in a group, and a results slide with a chart and a table.</summary>
    internal static List<(DeckSlideInfo Slide, IReadOnlyList<DeckObjectInfo> Objects)> SampleDeck() =>
    [
        (Slide(1, 256, "Agenda", appId: "agenda", section: "Intro", layout: "Title and Content"),
        [
            Shape(1, 256, 2, "Title 1", "placeholder", "Agenda", role: "title", placeholder: "ppPlaceholderTitle", minFont: 40),
            Shape(1, 256, 10, "Cards", "group", appId: "cards"),
            Shape(1, 256, 11, "Card A", "auto-shape", "Revenue grew 12%", role: "card", group: 10, minFont: 14, component: "card@1"),
            Shape(1, 256, 12, "Card B", "auto-shape", "Costs fell", role: "card", group: 10, minFont: 9, component: "card@1"),
        ]),
        (Slide(2, 260, "Q3 Results", section: "Results", tags: new Dictionary<string, string> { ["OWNER"] = "finance" }),
        [
            Shape(2, 260, 2, "Title 1", "placeholder", "Q3 Results", role: "title", placeholder: "ppPlaceholderTitle", minFont: 40),
            Shape(2, 260, 4, "Chart 3", "chart", appId: "q3-chart", left: 40, top: 120, width: 400, height: 300),
            Shape(2, 260, 5, "Table 4", "table", left: 460, top: 120, width: 460, height: 300, visible: false,
                tags: new Dictionary<string, string> { ["SOURCE"] = "q3.csv" }),
        ]),
    ];
}
