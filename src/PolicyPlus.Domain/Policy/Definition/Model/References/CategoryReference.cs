using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.References;

/// <summary>
///     Represents a reference to a category.
/// </summary>
public record struct CategoryReference
    : IEqualityOperators<CategoryReference, CategoryReference, bool>
{
    /// <summary>
    ///     Gets the category reference.
    /// </summary>
    public required string Ref { get; init; }
}