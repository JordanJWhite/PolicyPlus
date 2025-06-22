using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a single-line text box presentation element.
/// </summary>
public record TextBoxElement
    : IPresentationElement,
      IEqualityOperators<TextBoxElement, TextBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets the default value for the text box.
    /// </summary>
    public string? DefaultValue { get; init; }

    /// <summary>
    ///     Gets the reference ID to the text box.
    /// </summary>
    public required string RefId { get; init; }
}