using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a 64-bit decimal number element in a policy.
/// </summary>
public record LongDecimalElement
    : PolicyElementBase,
      IEqualityOperators<LongDecimalElement, LongDecimalElement, bool>
{
    /// <summary>
    ///     Gets whether this element is required.
    /// </summary>
    public bool Required { get; init; } = false;

    /// <summary>
    ///     Gets the minimum allowed value.
    /// </summary>
    public ulong MinValue { get; init; } = 0;

    /// <summary>
    ///     Gets the maximum allowed value.
    /// </summary>
    public ulong MaxValue { get; init; } = 9999;

    /// <summary>
    ///     Gets whether to store the value as text.
    /// </summary>
    public bool StoreAsText { get; init; } = false;

    /// <summary>
    ///     Gets whether this is a soft preference.
    /// </summary>
    public bool Soft { get; init; } = false;
}