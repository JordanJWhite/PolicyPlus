using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents an item in a value list.
/// </summary>
public record ValueItem
    : IEqualityOperators<ValueItem, ValueItem, bool>
{
    /// <summary>
    ///     Gets the value.
    /// </summary>
    public required IValue Value { get; init; }

    /// <summary>
    ///     Gets the optional registry key.
    /// </summary>
    public string? Key { get; init; }

    /// <summary>
    ///     Gets the registry value name.
    /// </summary>
    public required string ValueName { get; init; }
}