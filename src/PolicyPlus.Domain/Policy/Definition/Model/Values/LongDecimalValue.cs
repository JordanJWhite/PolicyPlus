using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents a long decimal (64-bit) value.
/// </summary>
public record struct LongDecimalValue
    : IValue,
      IEqualityOperators<LongDecimalValue, LongDecimalValue, bool>
{
    /// <summary>
    ///     Gets the long decimal value.
    /// </summary>
    public required ulong Value { get; init; }
}