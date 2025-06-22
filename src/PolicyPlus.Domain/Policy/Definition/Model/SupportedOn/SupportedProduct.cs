using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a supported product.
/// </summary>
public record SupportedProduct
    : IEqualityOperators<SupportedProduct, SupportedProduct, bool>
{
    /// <summary>
    ///     Gets the product name.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the major versions of this product.
    /// </summary>
    public IReadOnlyList<SupportedMajorVersion> MajorVersions { get; init; } = [];
}