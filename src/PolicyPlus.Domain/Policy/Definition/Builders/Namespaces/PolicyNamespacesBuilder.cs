using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Namespaces;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Namespaces;

/// <summary>
///     Builder interface for PolicyNamespaces.
/// </summary>
public interface IPolicyNamespacesBuilder : IBuilder<PolicyNamespaces>
{
    IPolicyNamespacesBuilder WithTarget(PolicyNamespaceAssociation                target);
    IPolicyNamespacesBuilder AddUsing(PolicyNamespaceAssociation                  usingNamespace);
    IPolicyNamespacesBuilder AddUsings(IEnumerable<PolicyNamespaceAssociation>    usingNamespaces);
    IPolicyNamespacesBuilder RemoveUsing(PolicyNamespaceAssociation               usingNamespace);
    IPolicyNamespacesBuilder RemoveUsings(IEnumerable<PolicyNamespaceAssociation> usingNamespaces);
    IPolicyNamespacesBuilder ClearUsings();
}

/// <summary>
///     Builder for PolicyNamespaces.
/// </summary>
public class PolicyNamespacesBuilder : BuilderBase<PolicyNamespacesBuilder, PolicyNamespaces>, IPolicyNamespacesBuilder
{
    private List<PolicyNamespaceAssociation>           _using = [];
    private IReadOnlyList<PolicyNamespaceAssociation>? _usingView;

    public PolicyNamespaceAssociation?               Target { get; private set; }
    public IReadOnlyList<PolicyNamespaceAssociation> Using  => _usingView ??= _using.AsReadOnly();

    public IPolicyNamespacesBuilder WithTarget(PolicyNamespaceAssociation target)
    {
        EnsureNotBuilt(nameof(WithTarget));
        Target = target;

        return this;
    }

    public IPolicyNamespacesBuilder AddUsing(PolicyNamespaceAssociation usingNamespace)
    {
        EnsureNotBuilt(nameof(AddUsing));
        _using.Add(usingNamespace);

        return this;
    }

    public IPolicyNamespacesBuilder AddUsings(IEnumerable<PolicyNamespaceAssociation> usingNamespaces)
    {
        ArgumentNullException.ThrowIfNull(usingNamespaces, nameof(usingNamespaces));
        EnsureNotBuilt(nameof(AddUsings));

        foreach (var usingNamespace in usingNamespaces)
        {
            if (usingNamespace.Equals(default))
                throw new ArgumentException("Using namespaces collection cannot contain default values.", nameof(usingNamespaces));

            _using.Add(usingNamespace);
        }

        return this;
    }

    public IPolicyNamespacesBuilder RemoveUsing(PolicyNamespaceAssociation usingNamespace)
    {
        // No need to check for null since it's a struct
        EnsureNotBuilt(nameof(RemoveUsing));

        _using.Remove(usingNamespace);

        return this;
    }

    public IPolicyNamespacesBuilder RemoveUsings(IEnumerable<PolicyNamespaceAssociation> usingNamespaces)
    {
        ArgumentNullException.ThrowIfNull(usingNamespaces, nameof(usingNamespaces));
        EnsureNotBuilt(nameof(RemoveUsings));

        foreach (var usingNamespace in usingNamespaces)
            _using.Remove(usingNamespace);

        return this;
    }

    public IPolicyNamespacesBuilder ClearUsings()
    {
        EnsureNotBuilt(nameof(ClearUsings));
        _using.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyNamespacesBuilder, PolicyNamespaces>(() => (nameof(Target),
                                                                                                                  Target is not null)
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
        _using     = [];
        _usingView = null;
    }
}