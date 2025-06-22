using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a reference item in a supported-on condition.
/// </summary>
public record struct SupportedOnReferenceItem
    : ISupportedOnItem,
      IEqualityOperators<SupportedOnReferenceItem, SupportedOnReferenceItem, bool>
{
    /// <summary>
    ///     Gets the reference.
    /// </summary>
    public required SupportedOnReference Reference { get; init; }
}