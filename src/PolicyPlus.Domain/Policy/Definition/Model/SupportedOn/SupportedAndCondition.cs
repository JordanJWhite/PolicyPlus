using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents an AND condition where all items must be true.
/// </summary>
public record struct SupportedAndCondition
    : ISupportedOnCondition,
      IEqualityOperators<SupportedAndCondition, SupportedAndCondition, bool>
{
    /// <summary>
    ///     Gets the condition items.
    /// </summary>
    public required IReadOnlyList<ISupportedOnItem> Items { get; init; }
}