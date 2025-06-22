using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents an item in an enumeration element.
/// </summary>
public record EnumerationItem
    : IEqualityOperators<EnumerationItem, EnumerationItem, bool>
{
    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the value for this item.
    /// </summary>
    public required IValue Value { get; init; }

    /// <summary>
    ///     Gets the optional value list for this item.
    /// </summary>

    public ValueList? ValueList { get; init; }
}