using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a range item in a supported-on condition.
/// </summary>
public record struct SupportedOnRangeItem
    : ISupportedOnItem,
      IEqualityOperators<SupportedOnRangeItem, SupportedOnRangeItem, bool>
{
    /// <summary>
    ///     Gets the range.
    /// </summary>
    public required SupportedOnRange Range { get; init; }
}