using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

namespace PolicyPlus.Domain.Policy.Definition.Builders;

/// <summary>
///     Builder interface for PolicyDefinitionResources.
/// </summary>
public interface IPolicyDefinitionResourcesBuilder : IBuilder<PolicyDefinitionResources>
{
    IPolicyDefinitionResourcesBuilder WithRevision(string                revision);
    IPolicyDefinitionResourcesBuilder WithSchemaVersion(string           schemaVersion);
    IPolicyDefinitionResourcesBuilder WithDisplayName(string             displayName);
    IPolicyDefinitionResourcesBuilder WithDescription(string             description);
    IPolicyDefinitionResourcesBuilder AddAnnotation(Annotation           annotation);
    IPolicyDefinitionResourcesBuilder AddString(LocalizedString          localizedString);
    IPolicyDefinitionResourcesBuilder AddPresentation(PolicyPresentation presentation);
}

/// <summary>
///     Builder for PolicyDefinitionResources.
/// </summary>
public class PolicyDefinitionResourcesBuilder
    : BuilderBase<PolicyDefinitionResourcesBuilder, PolicyDefinitionResources>, IPolicyDefinitionResourcesBuilder
{
    // Collections to hold annotations, strings, and presentations.
    private List<Annotation>         _annotations   = [];
    private List<LocalizedString>    _strings       = [];
    private List<PolicyPresentation> _presentations = [];

    // Views for read-only access to the collections.
    private IReadOnlyList<Annotation>?         _annotationsView;
    private IReadOnlyList<LocalizedString>?    _stringsView;
    private IReadOnlyList<PolicyPresentation>? _presentationsView;

    // Flags to track which properties have been set.
    private bool _descriptionSet;
    private bool _displayNameSet;
    private bool _revisionSet;
    private bool _schemaVersionSet;

    public string?                           Revision      { get; private set; }
    public string?                           SchemaVersion { get; private set; }
    public string?                           DisplayName   { get; private set; }
    public string?                           Description   { get; private set; }
    public IReadOnlyList<Annotation>         Annotations   => _annotationsView ??= _annotations.AsReadOnly();
    public IReadOnlyList<LocalizedString>    Strings       => _stringsView ??= _strings.AsReadOnly();
    public IReadOnlyList<PolicyPresentation> Presentations => _presentationsView ??= _presentations.AsReadOnly();

    public IPolicyDefinitionResourcesBuilder WithRevision(string revision)
    {
        ArgumentNullException.ThrowIfNull(revision, nameof(revision));
        EnsureNotBuilt(nameof(WithRevision));

        Revision     = revision;
        _revisionSet = true;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithSchemaVersion(string schemaVersion)
    {
        ArgumentNullException.ThrowIfNull(schemaVersion, nameof(schemaVersion));
        EnsureNotBuilt(nameof(WithSchemaVersion));

        SchemaVersion     = schemaVersion;
        _schemaVersionSet = true;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName     = displayName;
        _displayNameSet = true;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithDescription(string description)
    {
        ArgumentNullException.ThrowIfNull(description, nameof(description));
        EnsureNotBuilt(nameof(WithDescription));

        Description     = description;
        _descriptionSet = true;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddAnnotation(Annotation annotation)
    {
        ArgumentNullException.ThrowIfNull(annotation, nameof(annotation));
        EnsureNotBuilt(nameof(AddAnnotation));

        _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddString(LocalizedString localizedString)
    {
        ArgumentNullException.ThrowIfNull(localizedString, nameof(localizedString));
        EnsureNotBuilt(nameof(AddString));

        _strings.Add(localizedString);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddPresentation(PolicyPresentation presentation)
    {
        ArgumentNullException.ThrowIfNull(presentation, nameof(presentation));
        EnsureNotBuilt(nameof(AddPresentation));

        _presentations.Add(presentation);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyDefinitionResourcesBuilder, PolicyDefinitionResources>(
            () => (nameof(Revision), _revisionSet),
            () => (nameof(SchemaVersion), _schemaVersionSet),
            () => (nameof(DisplayName), _displayNameSet),
            () => (nameof(Description), _descriptionSet)
        );

    protected override PolicyDefinitionResources BuildCore() =>
        new()
        {
            Revision      = Revision!,
            SchemaVersion = SchemaVersion!,
            DisplayName   = DisplayName!,
            Description   = Description!,
            Annotations   = _annotations.AsReadOnly(),
            Strings       = _strings.AsReadOnly(),
            Presentations = _presentations.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Revision           = null;
        SchemaVersion      = null;
        DisplayName        = null;
        Description        = null;
        _annotations       = [];
        _strings           = [];
        _presentations     = [];
        _annotationsView   = null;
        _stringsView       = null;
        _presentationsView = null;
        _revisionSet       = false;
        _schemaVersionSet  = false;
        _displayNameSet    = false;
        _descriptionSet    = false;
    }
}