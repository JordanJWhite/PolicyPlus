using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.Namespaces;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;
using PolicyPlus.Domain.Policy.Definition.Model.References;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders;

/// <summary>
///     Builder interface for PolicyDefinitions.
/// </summary>
public interface IPolicyDefinitionsBuilder : IBuilder<PolicyDefinitions>
{
    IPolicyDefinitionsBuilder WithRevision(string                             revision);
    IPolicyDefinitionsBuilder WithSchemaVersion(string                        schemaVersion);
    IPolicyDefinitionsBuilder WithPolicyNamespaces(PolicyNamespaces           namespaces);
    IPolicyDefinitionsBuilder AddSupersededAdm(FileReference                  fileReference);
    IPolicyDefinitionsBuilder AddSupersededAdms(IEnumerable<FileReference>    fileReferences);
    IPolicyDefinitionsBuilder RemoveSupersededAdm(FileReference               fileReference);
    IPolicyDefinitionsBuilder RemoveSupersededAdms(IEnumerable<FileReference> fileReferences);
    IPolicyDefinitionsBuilder ClearSupersededAdms();
    IPolicyDefinitionsBuilder AddAnnotation(Annotation                  annotation);
    IPolicyDefinitionsBuilder AddAnnotations(IEnumerable<Annotation>    annotations);
    IPolicyDefinitionsBuilder RemoveAnnotation(Annotation               annotation);
    IPolicyDefinitionsBuilder RemoveAnnotations(IEnumerable<Annotation> annotations);
    IPolicyDefinitionsBuilder ClearAnnotations();
    IPolicyDefinitionsBuilder WithResources(LocalizationResourceReference resources);
    IPolicyDefinitionsBuilder WithSupportedOn(SupportedOnTable            supportedOn);
    IPolicyDefinitionsBuilder AddCategory(Category                        category);
    IPolicyDefinitionsBuilder AddCategories(IEnumerable<Category>         categories);
    IPolicyDefinitionsBuilder RemoveCategory(Category                     category);
    IPolicyDefinitionsBuilder RemoveCategories(IEnumerable<Category>      categories);
    IPolicyDefinitionsBuilder ClearCategories();
    IPolicyDefinitionsBuilder AddPolicy(PolicyDefinition                   policy);
    IPolicyDefinitionsBuilder AddPolicies(IEnumerable<PolicyDefinition>    policies);
    IPolicyDefinitionsBuilder RemovePolicy(PolicyDefinition                policy);
    IPolicyDefinitionsBuilder RemovePolicies(IEnumerable<PolicyDefinition> policies);
    IPolicyDefinitionsBuilder ClearPolicies();
}

/// <summary>
///     Builder for PolicyDefinitions.
/// </summary>
public class PolicyDefinitionsBuilder : BuilderBase<PolicyDefinitionsBuilder, PolicyDefinitions>, IPolicyDefinitionsBuilder
{
    // Collections to hold various elements of the policy definitions.
    private List<FileReference>    _supersededAdm = [];
    private List<Annotation>       _annotations   = [];
    private List<Category>         _categories    = [];
    private List<PolicyDefinition> _policies      = [];

    // Views for read-only access to the collections.
    private IReadOnlyList<FileReference>?    _supersededAdmView;
    private IReadOnlyList<Annotation>?       _annotationsView;
    private IReadOnlyList<Category>?         _categoriesView;
    private IReadOnlyList<PolicyDefinition>? _policiesView;

    // Flags to track which properties have been set.
    private bool _revisionSet;
    private bool _schemaVersionSet;
    private bool _policyNamespacesSet;
    private bool _resourcesSet;

    public string?                         Revision         { get; private set; }
    public string?                         SchemaVersion    { get; private set; }
    public PolicyNamespaces?               PolicyNamespaces { get; private set; }
    public IReadOnlyList<FileReference>    SupersededAdm    => _supersededAdmView ??= _supersededAdm.AsReadOnly();
    public IReadOnlyList<Annotation>       Annotations      => _annotationsView ??= _annotations.AsReadOnly();
    public LocalizationResourceReference?  Resources        { get; private set; }
    public SupportedOnTable?               SupportedOn      { get; private set; }
    public IReadOnlyList<Category>         Categories       => _categoriesView ??= _categories.AsReadOnly();
    public IReadOnlyList<PolicyDefinition> Policies         => _policiesView ??= _policies.AsReadOnly();

    public IPolicyDefinitionsBuilder WithRevision(string revision)
    {
        ArgumentNullException.ThrowIfNull(revision, nameof(revision));
        EnsureNotBuilt(nameof(WithRevision));

        Revision     = revision;
        _revisionSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder WithSchemaVersion(string schemaVersion)
    {
        ArgumentNullException.ThrowIfNull(schemaVersion, nameof(schemaVersion));
        EnsureNotBuilt(nameof(WithSchemaVersion));

        SchemaVersion     = schemaVersion;
        _schemaVersionSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder WithPolicyNamespaces(PolicyNamespaces namespaces)
    {
        ArgumentNullException.ThrowIfNull(namespaces, nameof(namespaces));
        EnsureNotBuilt(nameof(WithPolicyNamespaces));

        PolicyNamespaces     = namespaces;
        _policyNamespacesSet = true;

        return this;
    }

    public IPolicyDefinitionsBuilder AddSupersededAdm(FileReference fileReference)
    {
        EnsureNotBuilt(nameof(AddSupersededAdm));
        _supersededAdm.Add(fileReference);

        return this;
    }

    public IPolicyDefinitionsBuilder AddSupersededAdms(IEnumerable<FileReference> fileReferences)
    {
        ArgumentNullException.ThrowIfNull(fileReferences, nameof(fileReferences));
        EnsureNotBuilt(nameof(AddSupersededAdms));

        foreach (var fileReference in fileReferences)
            _supersededAdm.Add(fileReference);

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveSupersededAdm(FileReference fileReference)
    {
        EnsureNotBuilt(nameof(RemoveSupersededAdm));
        _supersededAdm.Remove(fileReference);

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveSupersededAdms(IEnumerable<FileReference> fileReferences)
    {
        ArgumentNullException.ThrowIfNull(fileReferences, nameof(fileReferences));
        EnsureNotBuilt(nameof(RemoveSupersededAdms));

        foreach (var fileReference in fileReferences)
            _supersededAdm.Remove(fileReference);

        return this;
    }

    public IPolicyDefinitionsBuilder ClearSupersededAdms()
    {
        EnsureNotBuilt(nameof(ClearSupersededAdms));
        _supersededAdm.Clear();

        return this;
    }

    public IPolicyDefinitionsBuilder AddAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(AddAnnotation));
        _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionsBuilder AddAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(AddAnnotations));

        foreach (var annotation in annotations)
            _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(RemoveAnnotation));
        _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(RemoveAnnotations));

        foreach (var annotation in annotations)
            _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionsBuilder ClearAnnotations()
    {
        EnsureNotBuilt(nameof(ClearAnnotations));
        _annotations.Clear();

        return this;
    }

    public IPolicyDefinitionsBuilder WithResources(LocalizationResourceReference resources)
    {
        EnsureNotBuilt(nameof(WithResources));

        Resources     = resources;
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

    public IPolicyDefinitionsBuilder AddCategories(IEnumerable<Category> categories)
    {
        ArgumentNullException.ThrowIfNull(categories, nameof(categories));
        EnsureNotBuilt(nameof(AddCategories));

        foreach (var category in categories)
        {
            if (category == null)
                throw new ArgumentNullException(nameof(categories), "Categories collection cannot contain null values.");

            _categories.Add(category);
        }

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveCategory(Category category)
    {
        ArgumentNullException.ThrowIfNull(category, nameof(category));
        EnsureNotBuilt(nameof(RemoveCategory));

        _categories.Remove(category);

        return this;
    }

    public IPolicyDefinitionsBuilder RemoveCategories(IEnumerable<Category> categories)
    {
        ArgumentNullException.ThrowIfNull(categories, nameof(categories));
        EnsureNotBuilt(nameof(RemoveCategories));

        foreach (var category in categories)
        {
            if (category == null)
                throw new ArgumentNullException(nameof(categories), "Categories collection cannot contain null values.");

            _categories.Remove(category);
        }

        return this;
    }

    public IPolicyDefinitionsBuilder ClearCategories()
    {
        EnsureNotBuilt(nameof(ClearCategories));
        _categories.Clear();
        _categoriesView = null;

        return this;
    }

    public IPolicyDefinitionsBuilder AddPolicy(PolicyDefinition policy)
    {
        ArgumentNullException.ThrowIfNull(policy, nameof(policy));
        EnsureNotBuilt(nameof(AddPolicy));

        _policies.Add(policy);

        return this;
    }

    public IPolicyDefinitionsBuilder AddPolicies(IEnumerable<PolicyDefinition> policies)
    {
        ArgumentNullException.ThrowIfNull(policies, nameof(policies));
        EnsureNotBuilt(nameof(AddPolicies));

        foreach (var policy in policies)
        {
            if (policy == null)
                throw new ArgumentNullException(nameof(policies), "Policies collection cannot contain null values.");

            _policies.Add(policy);
        }

        return this;
    }

    public IPolicyDefinitionsBuilder RemovePolicy(PolicyDefinition policy)
    {
        ArgumentNullException.ThrowIfNull(policy, nameof(policy));
        EnsureNotBuilt(nameof(RemovePolicy));

        _policies.Remove(policy);

        return this;
    }

    public IPolicyDefinitionsBuilder RemovePolicies(IEnumerable<PolicyDefinition> policies)
    {
        ArgumentNullException.ThrowIfNull(policies, nameof(policies));
        EnsureNotBuilt(nameof(RemovePolicies));

        foreach (var policy in policies)
        {
            if (policy == null)
                throw new ArgumentNullException(nameof(policies), "Policies collection cannot contain null values.");

            _policies.Remove(policy);
        }

        return this;
    }

    public IPolicyDefinitionsBuilder ClearPolicies()
    {
        EnsureNotBuilt(nameof(ClearPolicies));
        _policies.Clear();
        _policiesView = null;

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
            Revision         = Revision!,
            SchemaVersion    = SchemaVersion!,
            PolicyNamespaces = PolicyNamespaces!,
            SupersededAdm    = SupersededAdm,
            Annotations      = Annotations,
            Resources        = Resources!.Value,
            SupportedOn      = SupportedOn,
            Categories       = Categories,
            Policies         = Policies
        };

    protected override void ResetCore()
    {
        Revision             = null;
        SchemaVersion        = null;
        PolicyNamespaces     = null;
        Resources            = null;
        SupportedOn          = null;
        _supersededAdm       = [];
        _annotations         = [];
        _categories          = [];
        _policies            = [];
        _supersededAdmView   = null;
        _annotationsView     = null;
        _categoriesView      = null;
        _policiesView        = null;
        _revisionSet         = false;
        _schemaVersionSet    = false;
        _policyNamespacesSet = false;
        _resourcesSet        = false;
    }
}