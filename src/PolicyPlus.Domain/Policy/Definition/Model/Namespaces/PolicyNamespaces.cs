using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Namespaces;

/// <summary>
///     Represents policy namespaces configuration.
/// </summary>
public record PolicyNamespaces
    : IEqualityOperators<PolicyNamespaces, PolicyNamespaces, bool>
{
    /// <summary>
    ///     Gets the target namespace.
    /// </summary>
    public required PolicyNamespaceAssociation Target { get; init; }

    /// <summary>
    ///     Gets the using namespaces.
    /// </summary>
    public IReadOnlyList<PolicyNamespaceAssociation> Using { get; init; } = [];
}