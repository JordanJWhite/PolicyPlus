using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a list box presentation element.
/// </summary>
public record struct ListBoxElement
    : IPresentationElement,
      IEqualityOperators<ListBoxElement, ListBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text displayed above the list box.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets the reference ID to the list box.
    /// </summary>
    public required string RefId { get; init; }
}