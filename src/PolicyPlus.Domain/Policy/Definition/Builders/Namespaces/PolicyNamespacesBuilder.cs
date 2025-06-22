using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Namespaces;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Namespaces;

/// <summary>
///     Builder interface for PolicyNamespaces.
/// </summary>
public interface IPolicyNamespacesBuilder : IBuilder<PolicyNamespaces>
{
    IPolicyNamespacesBuilder WithTarget(PolicyNamespaceAssociation target);
    IPolicyNamespacesBuilder AddUsing(PolicyNamespaceAssociation   usingNamespace);
}

/// <summary>
///     Builder for PolicyNamespaces.
/// </summary>
public class PolicyNamespacesBuilder : BuilderBase<PolicyNamespacesBuilder, PolicyNamespaces>, IPolicyNamespacesBuilder
{
    private readonly List<PolicyNamespaceAssociation> _using = new();
    private          bool                             _targetSet;

    public PolicyNamespaceAssociation?               Target { get; private set; }
    public IReadOnlyList<PolicyNamespaceAssociation> Using  => _using.AsReadOnly();

    public IPolicyNamespacesBuilder WithTarget(PolicyNamespaceAssociation target)
    {
        ArgumentNullException.ThrowIfNull(target, nameof(target));
        EnsureNotBuilt(nameof(WithTarget));

        Target     = target;
        _targetSet = true;

        return this;
    }

    public IPolicyNamespacesBuilder AddUsing(PolicyNamespaceAssociation usingNamespace)
    {
        EnsureNotBuilt(nameof(AddUsing));

        _using.Add(usingNamespace);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyNamespacesBuilder, PolicyNamespaces>(() => (nameof(Target),
                                                                                                                  _targetSet)
        );

    protected override PolicyNamespaces BuildCore() =>
        new()
        {
            Target = Target!.Value,
            Using  = _using.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Target     = null;
        _targetSet = false;
        _using.Clear();
    }
}