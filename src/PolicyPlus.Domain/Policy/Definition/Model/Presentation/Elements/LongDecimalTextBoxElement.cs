using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a text box presentation element for a 64-bit decimal value.
/// </summary>
public record LongDecimalTextBoxElement
    : IPresentationElement,
      IEqualityOperators<LongDecimalTextBoxElement, LongDecimalTextBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets the default value.
    /// </summary>
    public ulong DefaultValue { get; init; } = 1;

    /// <summary>
    ///     Gets whether to show spin controls (up/down arrows).
    /// </summary>
    public bool Spin { get; init; } = true;

    /// <summary>
    ///     Gets the increment value for the spin controls.
    /// </summary>
    public uint SpinStep { get; init; } = 1;

    /// <summary>
    ///     Gets the reference ID to the long decimal text box.
    /// </summary>
    public required string RefId { get; init; }
}