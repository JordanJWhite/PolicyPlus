using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Represents a text box presentation element with a dropdown list of suggestions.
/// </summary>
public record struct ComboBoxElement()
    : IPresentationElement,
      IEqualityOperators<ComboBoxElement, ComboBoxElement, bool>
{
    /// <summary>
    ///     Gets the label text displayed next to the combo box.
    /// </summary>
    public required string Label { get; init; }

    /// <summary>
    ///     Gets the default text for the combo box.
    /// </summary>
    public string? DefaultValue { get; init; }

    /// <summary>
    ///     Gets the suggested values for the dropdown list.
    /// </summary>
    public IReadOnlyList<string> Suggestions { get; init; } = [];

    /// <summary>
    ///     Gets whether to disable alphabetical sorting of the suggestions.
    /// </summary>
    public bool NoSort { get; init; } = false;

    /// <summary>
    ///     Gets the reference ID to the combo box.
    /// </summary>
    public required string RefId { get; init; }
}