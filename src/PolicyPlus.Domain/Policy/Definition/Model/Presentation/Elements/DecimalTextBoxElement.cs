using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a decimal text box presentation element.
/// </summary>
public record DecimalTextBoxElement
    : IPresentationElement,
      IEqualityOperators<DecimalTextBoxElement, DecimalTextBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets the default value.
    /// </summary>
    public uint DefaultValue { get; init; } = 1;

    /// <summary>
    ///     Gets whether to show spin controls.
    /// </summary>
    public bool Spin { get; init; } = true;

    /// <summary>
    ///     Gets the spin step value.
    /// </summary>
    public uint SpinStep { get; init; } = 1;

    /// <summary>
    ///     Gets the reference ID to the decimal text box.
    /// </summary>
    public required string RefId { get; init; }
}