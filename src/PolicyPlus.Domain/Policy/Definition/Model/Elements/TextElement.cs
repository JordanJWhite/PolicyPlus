using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a text input element in a policy.
/// </summary>
public record TextElement
    : PolicyElementBase,
      IEqualityOperators<TextElement, TextElement, bool>
{
    /// <summary>
    ///     Gets whether this element is required.
    /// </summary>
    public bool Required { get; init; } = false;

    /// <summary>
    ///     Gets the maximum text length.
    /// </summary>
    public uint MaxLength { get; init; } = 1023;

    /// <summary>
    ///     Gets whether environment variables should be expanded.
    /// </summary>
    public bool Expandable { get; init; } = false;

    /// <summary>
    ///     Gets whether this is a soft preference.
    /// </summary>
    public bool Soft { get; init; } = false;
}