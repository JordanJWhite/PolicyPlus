using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents an OR condition where at least one item must be true.
/// </summary>
public record struct SupportedOrCondition
    : ISupportedOnCondition,
      IEqualityOperators<SupportedOrCondition, SupportedOrCondition, bool>
{
    /// <summary>
    ///     Gets the condition items.
    /// </summary>
    public required IReadOnlyList<ISupportedOnItem> Items { get; init; }
}