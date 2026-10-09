namespace Sbroenne.PowerPointMcp.Core.Design;

/// <summary>A component style in a profile and its instances in the deck.</summary>
public sealed class DesignComponentInfo
{
    /// <summary>Component name.</summary>
    public required string Name { get; init; }

    /// <summary>Version the profile defines (null when the profile has no such component).</summary>
    public string? ProfileVersion { get; init; }

    /// <summary>Number of tagged instances in the deck.</summary>
    public int Instances { get; init; }

    /// <summary>Instances whose version differs from the profile's.</summary>
    public int Outdated { get; init; }

    /// <summary>Versions found in the deck with counts, e.g. "1": 4.</summary>
    public IReadOnlyDictionary<string, int>? DeckVersions { get; init; }

    /// <summary>Slides that hold instances.</summary>
    public IReadOnlyList<int>? Slides { get; init; }
}

/// <summary>Result of a design command. Success == true implies ErrorMessage is null (Rule 1).</summary>
public sealed class DesignOperationResult
{
    /// <summary>Whether the operation succeeded.</summary>
    public bool Success { get; init; }

    /// <summary>Error message when Success is false.</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>Profiles available (built-in, saved, and the deck theme).</summary>
    public IReadOnlyList<ProfileEntry>? Profiles { get; init; }

    /// <summary>Profile name the result is about.</summary>
    public string? ProfileName { get; init; }

    /// <summary>Profile definition JSON (pptmcp.design-profile/1).</summary>
    public string? ProfileJson { get; init; }

    /// <summary>Resolved profile JSON: inheritance applied and every token as a value.</summary>
    public string? ResolvedJson { get; init; }

    /// <summary>Validation errors.</summary>
    public IReadOnlyList<string>? Errors { get; init; }

    /// <summary>File written or read.</summary>
    public string? Path { get; init; }

    /// <summary>Planned or applied changes (paginated by limit).</summary>
    public IReadOnlyList<DesignChange>? Changes { get; init; }

    /// <summary>Total number of planned changes.</summary>
    public int? ChangeCount { get; init; }

    /// <summary>Planned changes per property (font-name, font-size, text-color, fill, ...).</summary>
    public IReadOnlyDictionary<string, int>? ChangesByProperty { get; init; }

    /// <summary>Whether changes were written (false for a dry run).</summary>
    public bool? Applied { get; init; }

    /// <summary>Components defined by the profile and found in the deck.</summary>
    public IReadOnlyList<DesignComponentInfo>? Components { get; init; }

    /// <summary>How the result was derived, what was kept, and follow-up advice.</summary>
    public IReadOnlyList<string>? Warnings { get; init; }
}
