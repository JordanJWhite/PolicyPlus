using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

/// <summary>
///     Represents a complex supported-on definition.
/// </summary>
public record SupportedOnDefinition
    : IEqualityOperators<SupportedOnDefinition, SupportedOnDefinition, bool>
{
    /// <summary>
    ///     Gets the definition name.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the condition for this definition.
    /// </summary>
    public ISupportedOnCondition? Condition { get; init; }
}