using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

/// <summary>
///     Represents a reference to localization resources.
/// </summary>
public record struct LocalizationResourceReference()
    : IEqualityOperators<LocalizationResourceReference, LocalizationResourceReference, bool>
{
    /// <summary>
    ///     Gets the minimum required revision.
    /// </summary>
    public required string MinRequiredRevision { get; init; }

    /// <summary>
    ///     Gets the fallback culture.
    /// </summary>
    public string FallbackCulture { get; init; } = "en-US";
}