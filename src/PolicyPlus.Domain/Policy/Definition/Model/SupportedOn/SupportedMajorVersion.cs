using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a major version of a supported product.
/// </summary>
public record SupportedMajorVersion
    : IEqualityOperators<SupportedMajorVersion, SupportedMajorVersion, bool>
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

    /// <summary>
    ///     Gets the minor versions of this major version.
    /// </summary>
    public IReadOnlyList<SupportedMinorVersion> MinorVersions { get; init; } = [];
}