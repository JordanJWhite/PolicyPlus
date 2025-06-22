using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents a decimal value.
/// </summary>
public record struct DecimalValue
    : IValue,
      IEqualityOperators<DecimalValue, DecimalValue, bool>
{
    /// <summary>
    ///     Gets the decimal value.
    /// </summary>
    public required uint Value { get; init; }
}