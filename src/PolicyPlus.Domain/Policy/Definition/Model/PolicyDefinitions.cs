using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.Namespaces;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;
using PolicyPlus.Domain.Policy.Definition.Model.References;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Model;

/// <summary>
///     Represents the policy definitions.
/// </summary>
public record PolicyDefinitions : IEqualityOperators<PolicyDefinitions, PolicyDefinitions, bool>
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
    ///     Gets the policy namespaces.
    /// </summary>
    public required PolicyNamespaces PolicyNamespaces { get; init; }

    /// <summary>
    ///     Gets the superseded ADM files.
    /// </summary>
    public IReadOnlyList<FileReference> SupersededAdm { get; init; } = [];

    /// <summary>
    ///     Gets the application-specific annotations.
    /// </summary>
    public IReadOnlyList<Annotation> Annotations { get; init; } = [];

    /// <summary>
    ///     Gets the localization resource reference.
    /// </summary>
    public required LocalizationResourceReference Resources { get; init; }

    /// <summary>
    ///     Gets the supported-on table.
    /// </summary>
    public SupportedOnTable? SupportedOn { get; init; }

    /// <summary>
    ///     Gets the categories.
    /// </summary>
    public IReadOnlyList<Category> Categories { get; init; } = [];

    /// <summary>
    ///     Gets the policy definitions.
    /// </summary>
    public IReadOnlyList<PolicyDefinition> Policies { get; init; } = [];
}