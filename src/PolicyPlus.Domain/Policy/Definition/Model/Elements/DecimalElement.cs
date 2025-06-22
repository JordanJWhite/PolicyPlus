using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a decimal number element in a policy.
/// </summary>
public record DecimalElement
    : PolicyElementBase,
      IEqualityOperators<DecimalElement, DecimalElement, bool>
{
    /// <summary>
    ///     Gets whether this element is required.
    /// </summary>
    public bool Required { get; init; } = false;

    /// <summary>
    ///     Gets the minimum allowed value.
    /// </summary>
    public uint MinValue { get; init; } = 0;

    /// <summary>
    ///     Gets the maximum allowed value.
    /// </summary>
    public uint MaxValue { get; init; } = 9999;

    /// <summary>
    ///     Gets whether to store the value as text.
    /// </summary>
    public bool StoreAsText { get; init; } = false;

    /// <summary>
    ///     Gets whether this is a soft preference.
    /// </summary>
    public bool Soft { get; init; } = false;
}