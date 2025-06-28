using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Namespaces;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Namespaces;

/// <summary>
///     Builder interface for PolicyNamespaceAssociation.
/// </summary>
public interface IPolicyNamespaceAssociationBuilder : IBuilder<PolicyNamespaceAssociation>
{
    IPolicyNamespaceAssociationBuilder WithPrefix(string    prefix);
    IPolicyNamespaceAssociationBuilder WithNamespace(string namespaceUri);
}

/// <summary>
///     Builder for PolicyNamespaceAssociation.
/// </summary>
public class PolicyNamespaceAssociationBuilder
    : BuilderBase<PolicyNamespaceAssociationBuilder, PolicyNamespaceAssociation>, IPolicyNamespaceAssociationBuilder
{
    public string? Prefix    { get; private set; }
    public string? Namespace { get; private set; }

    public IPolicyNamespaceAssociationBuilder WithPrefix(string prefix)
    {
        ArgumentNullException.ThrowIfNull(prefix, nameof(prefix));
        EnsureNotBuilt(nameof(WithPrefix));

        Prefix = prefix;

        return this;
    }

    public IPolicyNamespaceAssociationBuilder WithNamespace(string namespaceUri)
    {
        ArgumentNullException.ThrowIfNull(namespaceUri, nameof(namespaceUri));
        EnsureNotBuilt(nameof(WithNamespace));

        Namespace = namespaceUri;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyNamespaceAssociationBuilder, PolicyNamespaceAssociation>(
            () => (nameof(Prefix), Prefix is not null),
            () => (nameof(Namespace), Namespace is not null)
        );

    protected override PolicyNamespaceAssociation BuildCore() =>
        new()
        {
            Prefix    = Prefix!,
            Namespace = Namespace!
        };

    protected override void ResetCore()
    {
        Prefix    = null;
        Namespace = null;
    }
}