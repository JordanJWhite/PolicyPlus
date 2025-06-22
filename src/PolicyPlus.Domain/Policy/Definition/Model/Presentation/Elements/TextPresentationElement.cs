using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a text presentation element.
/// </summary>
public record struct TextPresentationElement
    : IPresentationElement,
      IEqualityOperators<TextPresentationElement, TextPresentationElement, bool>
{
    /// <summary>
    ///     Gets the text content.
    /// </summary>
    public required string Text { get; init; }

    /// <summary>
    ///     Gets the reference ID to the text.
    /// </summary>
    public required string RefId { get; init; }
}