using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a multi-line text box presentation element.
/// </summary>
public record MultiTextBoxElement
    : IPresentationElement,
      IEqualityOperators<MultiTextBoxElement, MultiTextBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text displayed with the text box.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets whether to display the text box in a separate dialog window.
    /// </summary>
    public bool ShowAsDialog { get; init; } = false;

    /// <summary>
    ///     Gets the default height of the text box in lines.
    /// </summary>
    public uint DefaultHeight { get; init; } = 3;

    /// <summary>
    ///     Gets the reference ID to the multi-line text box.
    /// </summary>
    public required string RefId { get; init; }
}