using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents a string value.
/// </summary>
public record struct StringValue
    : IValue,
      IEqualityOperators<StringValue, StringValue, bool>
{
    /// <summary>
    ///     Gets the string value.
    /// </summary>
    public required string Value { get; init; }
}