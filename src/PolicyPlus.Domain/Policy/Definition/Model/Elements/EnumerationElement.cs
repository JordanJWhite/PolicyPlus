using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents an enumeration (dropdown) element in a policy.
/// </summary>
public record EnumerationElement
    : PolicyElementBase,
      IEqualityOperators<EnumerationElement, EnumerationElement, bool>
{
    /// <summary>
    ///     Gets whether this element is required.
    /// </summary>
    public bool Required { get; init; } = false;

    /// <summary>
    ///     Gets the enumeration items.
    /// </summary>
    public required IReadOnlyList<EnumerationItem> Items { get; init; }
}