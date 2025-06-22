using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a supported version range.
/// </summary>
public record struct SupportedOnRange
    : IEqualityOperators<SupportedOnRange, SupportedOnRange, bool>
{
    /// <summary>
    ///     Gets the reference identifier.
    /// </summary>
    public required string Ref { get; init; }

    /// <summary>
    ///     Gets the optional minimum version index.
    /// </summary>
    public uint? MinVersionIndex { get; init; }

    /// <summary>
    ///     Gets the optional maximum version index.
    /// </summary>
    public uint? MaxVersionIndex { get; init; }
}