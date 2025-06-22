using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.References;

namespace PolicyPlus.Domain.Policy.Definition.Model.Core;

/// <summary>
///     Represents a category for grouping policy definitions.
/// </summary>
public record Category
    : IEqualityOperators<Category, Category, bool>
{
    /// <summary>
    ///     Gets the category name.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the application-specific annotations.
    /// </summary>
    public IReadOnlyList<Annotation> Annotations { get; init; } = [];

    /// <summary>
    ///     Gets the optional explanation text reference.
    /// </summary>
    public string? ExplainText { get; init; }

    /// <summary>
    ///     Gets the parent category reference.
    /// </summary>
    public CategoryReference? ParentCategory { get; init; }

    /// <summary>
    ///     Gets the optional keywords.
    /// </summary>
    public string? Keywords { get; init; }

    /// <summary>
    ///     Gets the see also references.
    /// </summary>
    public IReadOnlyList<string> SeeAlso { get; init; } = [];
}