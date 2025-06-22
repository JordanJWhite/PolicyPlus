using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a checkbox presentation element.
/// </summary>
public record struct CheckBoxElement()
    : IPresentationElement,
      IEqualityOperators<CheckBoxElement, CheckBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets whether the checkbox is checked by default.
    /// </summary>
    public bool DefaultChecked { get; init; } = false;

    /// <summary>
    ///     Gets the reference ID to the checkbox.
    /// </summary>
    public required string RefId { get; init; }
}