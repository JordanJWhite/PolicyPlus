using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a list element in a policy.
/// </summary>
public record ListElement
    : PolicyElementBase,
      IEqualityOperators<ListElement, ListElement, bool>
{
    /// <summary>
    ///     Gets the value prefix for list items.
    /// </summary>
    public string? ValuePrefix { get; init; }

    /// <summary>
    ///     Gets whether the list is additive.
    /// </summary>
    public bool Additive { get; init; } = false;

    /// <summary>
    ///     Gets whether environment variables should be expanded.
    /// </summary>
    public bool Expandable { get; init; } = false;

    /// <summary>
    ///     Gets whether to use explicit value names.
    /// </summary>
    public bool ExplicitValue { get; init; } = false;
}