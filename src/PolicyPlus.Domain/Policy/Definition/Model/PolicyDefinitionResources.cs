using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

namespace PolicyPlus.Domain.Policy.Definition.Model;

/// <summary>
///     Represents policy definition resources.
/// </summary>
public record PolicyDefinitionResources : IEqualityOperators<PolicyDefinitionResources, PolicyDefinitionResources, bool>
{
    /// <summary>
    ///     Gets the revision version.
    /// </summary>
    public required string Revision { get; init; }

    /// <summary>
    ///     Gets the schema version.
    /// </summary>
    public required string SchemaVersion { get; init; }

    /// <summary>
    ///     Gets the display name.
    /// </summary>
    public required string DisplayName { get; init; }

    /// <summary>
    ///     Gets the description.
    /// </summary>
    public required string Description { get; init; }

    /// <summary>
    ///     Gets the application-specific annotations.
    /// </summary>
    public IReadOnlyList<Annotation> Annotations { get; init; } = [];

    /// <summary>
    ///     Gets the localized strings.
    /// </summary>
    public IReadOnlyList<LocalizedString> Strings { get; init; } = [];

    /// <summary>
    ///     Gets the policy presentations.
    /// </summary>
    public IReadOnlyList<PolicyPresentation> Presentations { get; init; } = [];
}