using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a minor version of a supported product.
/// </summary>
public record SupportedMinorVersion
    : IEqualityOperators<SupportedMinorVersion, SupportedMinorVersion, bool>
{
    /// <summary>
    ///     Gets the version name.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the version index for ordering.
    /// </summary>
    public required uint VersionIndex { get; init; }
}