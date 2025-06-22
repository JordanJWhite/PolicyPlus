using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a multi-line text element in a policy.
/// </summary>
public record MultiTextElement
    : PolicyElementBase,
      IEqualityOperators<MultiTextElement, MultiTextElement, bool>
{
    /// <summary>
    ///     Gets whether this element is required.
    /// </summary>
    public bool Required { get; init; } = false;

    /// <summary>
    ///     Gets the maximum text length per line.
    /// </summary>
    public uint MaxLength { get; init; } = 1023;

    /// <summary>
    ///     Gets the maximum number of strings (0 = unlimited).
    /// </summary>
    public uint MaxStrings { get; init; } = 0;

    /// <summary>
    ///     Gets whether this is a soft preference.
    /// </summary>
    public bool Soft { get; init; } = false;
}