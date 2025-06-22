using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a reference to a supported product definition.
/// </summary>
public record struct SupportedOnReference
    : IEqualityOperators<SupportedOnReference, SupportedOnReference, bool>
{
    /// <summary>
    ///     Gets the reference identifier.
    /// </summary>
    public required string Ref { get; init; }
}