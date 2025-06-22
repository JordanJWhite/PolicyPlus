using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents a list of values.
/// </summary>
public record struct ValueList
    : IEqualityOperators<ValueList, ValueList, bool>
{
    /// <summary>
    ///     Gets the list of value items.
    /// </summary>
    public required IReadOnlyList<ValueItem> Items { get; init; }

    /// <summary>
    ///     Gets the optional default registry key.
    /// </summary>
    public string? DefaultKey { get; init; }
}