using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a dropdown list presentation element.
/// </summary>
public record DropdownListElement
    : IPresentationElement,
      IEqualityOperators<DropdownListElement, DropdownListElement, bool>
{
    /// <summary>
    ///     Gets the label text.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets whether to disable sorting.
    /// </summary>
    public bool NoSort { get; init; } = false;

    /// <summary>
    ///     Gets the default item index.
    /// </summary>
    public uint? DefaultItem { get; init; }

    /// <summary>
    ///     Gets the reference ID to the drop down.
    /// </summary>
    public required string RefId { get; init; }
}