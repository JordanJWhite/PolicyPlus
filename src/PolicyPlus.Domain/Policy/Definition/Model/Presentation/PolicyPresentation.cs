using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Model.Presentation;

/// <summary>
///     Represents a policy presentation definition.
/// </summary>
public record PolicyPresentation
    : IEqualityOperators<PolicyPresentation, PolicyPresentation, bool>
{
    /// <summary>
    ///     Gets the presentation identifier.
    /// </summary>
    public required string Id { get; init; }

    /// <summary>
    ///     Gets the presentation elements.
    /// </summary>
    public IReadOnlyList<IPresentationElement> Elements { get; init; } = [];
}