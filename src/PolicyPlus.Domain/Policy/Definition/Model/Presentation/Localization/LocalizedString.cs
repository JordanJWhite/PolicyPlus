using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

/// <summary>
///     Represents a localized string.
/// </summary>
public record struct LocalizedString
    : IEqualityOperators<LocalizedString, LocalizedString, bool>
{
    /// <summary>
    ///     Gets the string identifier.
    /// </summary>
    public required string Id { get; init; }

    /// <summary>
    ///     Gets the string value.
    /// </summary>
    public required string Value { get; init; }
}