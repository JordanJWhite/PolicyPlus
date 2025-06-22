using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Namespaces;

/// <summary>
///     Represents a policy namespace association.
/// </summary>
public record struct PolicyNamespaceAssociation
    : IEqualityOperators<PolicyNamespaceAssociation, PolicyNamespaceAssociation, bool>
{
    /// <summary>
    ///     Gets the namespace prefix.
    /// </summary>
    public required string Prefix { get; init; }

    /// <summary>
    ///     Gets the namespace URI.
    /// </summary>
    public required string Namespace { get; init; }
}