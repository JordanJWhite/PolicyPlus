using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;
using PolicyPlus.Domain.Policy.Definition.Model.References;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Model.Core;

/// <summary>
///     Represents a policy definition.
/// </summary>
public record PolicyDefinition
    : IEqualityOperators<PolicyDefinition, PolicyDefinition, bool>
{
    /// <summary>
    ///     Gets the policy name.
    /// </summary>
    public required string Name { get; init; }

    /// <summary>
    ///     Gets the policy class.
    /// </summary>
    public required PolicyClass Class { get; init; }

    /// <summary>
    ///     Gets the application-specific annotations.
    /// </summary>
    public IReadOnlyList<Annotation> Annotations { get; init; } = [];

    /// <summary>
    ///     Gets the display name reference.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the optional explanation text reference.
    /// </summary>
    public string? ExplainText { get; init; }

    /// <summary>
    ///     Gets the registry key.
    /// </summary>
    public required string Key { get; init; }

    /// <summary>
    ///     Gets the optional registry value name.
    /// </summary>
    public string? ValueName { get; init; }

    /// <summary>
    ///     Gets the parent category reference.
    /// </summary>
    public CategoryReference? ParentCategory { get; init; }

    /// <summary>
    ///     Gets the supported on reference.
    /// </summary>
    public required SupportedOnReference SupportedOn { get; init; }

    /// <summary>
    ///     Gets the optional presentation reference.
    /// </summary>
    public string? Presentation { get; init; }

    /// <summary>
    ///     Gets the value when enabled.
    /// </summary>
    public IValue? EnabledValue { get; init; }

    /// <summary>
    ///     Gets the value when disabled.
    /// </summary>
    public IValue? DisabledValue { get; init; }

    /// <summary>
    ///     Gets the value list when enabled.
    /// </summary>
    public ValueList? EnabledList { get; init; }

    /// <summary>
    ///     Gets the value list when disabled.
    /// </summary>
    public ValueList? DisabledList { get; init; }

    /// <summary>
    ///     Gets the policy elements.
    /// </summary>
    public IReadOnlyList<PolicyElementBase> Elements { get; init; } = [];

    /// <summary>
    ///     Gets the optional keywords.
    /// </summary>
    public string? Keywords { get; init; }

    /// <summary>
    ///     Gets the see also references.
    /// </summary>
    public IReadOnlyList<string> SeeAlso { get; init; } = [];
}