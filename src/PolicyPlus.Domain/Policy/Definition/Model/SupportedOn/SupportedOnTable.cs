using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents the supported-on table containing products and definitions.
/// </summary>
public record struct SupportedOnTable()
    : IEqualityOperators<SupportedOnTable, SupportedOnTable, bool>
{
    /// <summary>
    ///     Gets the supported products.
    /// </summary>
    public IReadOnlyList<SupportedProduct> Products { get; init; } = [];

    /// <summary>
    ///     Gets the supported-on definitions.
    /// </summary>
    public IReadOnlyList<SupportedOnDefinition> Definitions { get; init; } = [];
}