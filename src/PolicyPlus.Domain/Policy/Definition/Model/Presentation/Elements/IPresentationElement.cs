namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

/// <summary>
///     Interface for presentation elements.
/// </summary>
public interface IPresentationElement
{
    /// <summary>
    ///     Gets the reference ID to the policy element.
    /// </summary>
    public string RefId { get; init; }
}