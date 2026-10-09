namespace Sbroenne.PowerPointMcp.Core.Review;

/// <summary>Description of one validation rule.</summary>
public sealed class ValidationRule
{
    /// <summary>Rule code.</summary>
    public required string Code { get; init; }

    /// <summary>deterministic or heuristic.</summary>
    public required string Certainty { get; init; }

    /// <summary>How the rule detects problems.</summary>
    public required string Method { get; init; }

    /// <summary>What the rule checks.</summary>
    public required string Description { get; init; }

    /// <summary>Whether review repair fixes it automatically.</summary>
    public required bool AutoRepairable { get; init; }

    /// <summary>Whether repair includes it when no codes are named.</summary>
    public required bool RepairedByDefault { get; init; }
}
