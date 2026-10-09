namespace Sbroenne.PowerPointMcp.Core.Deck;

/// <summary>Semantic roles shared by inspection, selectors, composition, and design profiles.</summary>
public static class DeckRoles
{
    /// <summary>Tag holding a persistent application identifier.</summary>
    public const string IdTag = "PPTMCP_ID";

    /// <summary>Tag holding an explicit semantic role.</summary>
    public const string RoleTag = "PPTMCP_ROLE";

    /// <summary>Tag holding component identity as "name@version".</summary>
    public const string ComponentTag = "PPTMCP_COMPONENT";

    /// <summary>Tag holding the data binding of a table or chart (source file and range).</summary>
    public const string BindingTag = "PPTMCP_BINDING";

    /// <summary>Tag holding a diagram node identifier.</summary>
    public const string NodeTag = "PPTMCP_NODE";

    /// <summary>Maps a native PpPlaceholderType member name to a role.</summary>
    public static string? FromPlaceholder(string? placeholderType) => placeholderType switch
    {
        "ppPlaceholderTitle" or "ppPlaceholderCenterTitle" or "ppPlaceholderVerticalTitle" => "title",
        "ppPlaceholderSubtitle" => "subtitle",
        "ppPlaceholderBody" or "ppPlaceholderVerticalBody" => "body",
        "ppPlaceholderFooter" => "footer",
        "ppPlaceholderSlideNumber" => "slide-number",
        "ppPlaceholderDate" => "date",
        "ppPlaceholderHeader" => "header",
        "ppPlaceholderPicture" or "ppPlaceholderBitmap" => "picture",
        "ppPlaceholderChart" => "chart",
        "ppPlaceholderTable" => "table",
        "ppPlaceholderObject" or "ppPlaceholderVerticalObject" => "object",
        "ppPlaceholderMediaClip" => "media",
        "ppPlaceholderOrgChart" => "diagram",
        null => null,
        _ => "other",
    };
}
