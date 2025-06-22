using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Core;
using PolicyPlus.Domain.Policy.Definition.Namespaces;
using PolicyPlus.Domain.Policy.Definition.Presentation.Localization;
using PolicyPlus.Domain.Policy.Definition.References;
using PolicyPlus.Domain.Policy.Definition.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders;

/// <summary>
///     Builder interface for PolicyDefinitions.
/// </summary>
public interface IPolicyDefinitionsBuilder : IBuilder<PolicyDefinitions>
{
    IPolicyDefinitionsBuilder WithRevision(string revision);
    IPolicyDefinitionsBuilder WithSchemaVersion(string schemaVersion);
    IPolicyDefinitionsBuilder WithPolicyNamespaces(PolicyNamespaces namespaces);
    IPolicyDefinitionsBuilder AddSupersededAdm(FileReference fileReference);
    IPolicyDefinitionsBuilder AddAnnotation(Annotation annotation);
    IPolicyDefinitionsBuilder WithResources(LocalizationResourceReference resources);
    IPolicyDefinitionsBuilder WithSupportedOn(SupportedOnTable supportedOn);
    IPolicyDefinitionsBuilder AddCategory(Category category);
    IPolicyDefinitionsBuilder AddPolicy(PolicyDefinition policy);
}

/// <summary>
///     Builder for PolicyDefinitions.
/// </summary>
public class PolicyDefinitionsBuilder : BuilderBase<PolicyDefinitionsBuilder, PolicyDefinitions>, IPolicyDefinitionsBuilder
{
    private readonly List<Annotation> _annotations = new();
    private readonly List<Category> _categories = new();
    private readonly List<PolicyDefinition> _policies = new();
    private readonly List<FileReference> _supersededAdm = new();
    
    private bool _revisionSet;
    private bool _schemaVersionSet;
    private bool _policyNamespacesSet;
    private bool _resourcesSet;

    public string? Revision { get; private set; }
    public string? SchemaVersion { get; private set; }
    public PolicyNamespaces? PolicyNamespaces { get; private set; }
    public IReadOnlyList<FileReference> SupersededAdm => _supersededAdm.AsReadOnly();
    public IReadOnlyList<Annotation> Annotations => _annotations.AsReadOnly();
    public LocalizationResourceReference? Resources { get; private set; }
    public SupportedOnTable? SupportedOn { get; private set; }
    public IReadOnlyList<Category> Categories => _categories.AsReadOnly();
    public IReadOnlyList<PolicyDefinition> Policies => _policies.AsReadOnly();

    public IPolicyDefinitionsBuilder WithRevision(string revision)
    {
        ArgumentNullException.ThrowIfNull(revision, nameof(revision));
        EnsureNotBuilt(nameof(WithRevision));
        
        Revision = revision;
        _revisionSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder WithSchemaVersion(string schemaVersion)
    {
        ArgumentNullException.ThrowIfNull(schemaVersion, nameof(schemaVersion));
        EnsureNotBuilt(nameof(WithSchemaVersion));
        
        SchemaVersion = schemaVersion;
        _schemaVersionSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder WithPolicyNamespaces(PolicyNamespaces namespaces)
    {
        ArgumentNullException.ThrowIfNull(namespaces, nameof(namespaces));
        EnsureNotBuilt(nameof(WithPolicyNamespaces));
        
        PolicyNamespaces = namespaces;
        _policyNamespacesSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder AddSupersededAdm(FileReference fileReference)
    {
        EnsureNotBuilt(nameof(AddSupersededAdm));
        _supersededAdm.Add(fileReference);

        return this;
    }

    public IPolicyDefinitionsBuilder AddAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(AddAnnotation));
        _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionsBuilder WithResources(LocalizationResourceReference resources)
    {
        EnsureNotBuilt(nameof(WithResources));
        
        Resources = resources;
        _resourcesSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder WithSupportedOn(SupportedOnTable supportedOn)
    {
        EnsureNotBuilt(nameof(WithSupportedOn));
        SupportedOn = supportedOn;

        return this;
    }

    public IPolicyDefinitionsBuilder AddCategory(Category category)
    {
        ArgumentNullException.ThrowIfNull(category, nameof(category));
        EnsureNotBuilt(nameof(AddCategory));
        
        _categories.Add(category);

        return this;
    }

    public IPolicyDefinitionsBuilder AddPolicy(PolicyDefinition policy)
    {
        ArgumentNullException.ThrowIfNull(policy, nameof(policy));
        EnsureNotBuilt(nameof(AddPolicy));
        
        _policies.Add(policy);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyDefinitionsBuilder, PolicyDefinitions>(
            () => (nameof(Revision), _revisionSet),
            () => (nameof(SchemaVersion), _schemaVersionSet),
            () => (nameof(PolicyNamespaces), _policyNamespacesSet),
            () => (nameof(Resources), _resourcesSet)
        );

    protected override PolicyDefinitions BuildCore() =>
        new()
        {
            Revision = Revision!,
            SchemaVersion = SchemaVersion!,
            PolicyNamespaces = PolicyNamespaces!,
            SupersededAdm = _supersededAdm.AsReadOnly(),
            Annotations = _annotations.AsReadOnly(),
            Resources = Resources!.Value,
            SupportedOn = SupportedOn,
            Categories = _categories.AsReadOnly(),
            Policies = _policies.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Revision = null;
        SchemaVersion = null;
        PolicyNamespaces = null;
        _supersededAdm.Clear();
        _annotations.Clear();
        Resources = null;
        SupportedOn = null;
        _categories.Clear();
        _policies.Clear();
        _revisionSet = false;
        _schemaVersionSet = false;
        _policyNamespacesSet = false;
        _resourcesSet = false;
    }
}