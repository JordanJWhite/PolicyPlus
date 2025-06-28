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
    IPolicyDefinitionResourcesBuilder WithRevision(string                       revision);
    IPolicyDefinitionResourcesBuilder WithSchemaVersion(string                  schemaVersion);
    IPolicyDefinitionResourcesBuilder WithDisplayName(string                    displayName);
    IPolicyDefinitionResourcesBuilder WithDescription(string                    description);
    IPolicyDefinitionResourcesBuilder AddAnnotation(Annotation                  annotation);
    IPolicyDefinitionResourcesBuilder AddAnnotations(IEnumerable<Annotation>    annotations);
    IPolicyDefinitionResourcesBuilder RemoveAnnotation(Annotation               annotation);
    IPolicyDefinitionResourcesBuilder RemoveAnnotations(IEnumerable<Annotation> annotations);
    IPolicyDefinitionResourcesBuilder ClearAnnotations();
    IPolicyDefinitionResourcesBuilder AddString(LocalizedString                  localizedString);
    IPolicyDefinitionResourcesBuilder AddStrings(IEnumerable<LocalizedString>    localizedStrings);
    IPolicyDefinitionResourcesBuilder RemoveString(LocalizedString               localizedString);
    IPolicyDefinitionResourcesBuilder RemoveStrings(IEnumerable<LocalizedString> localizedStrings);
    IPolicyDefinitionResourcesBuilder ClearStrings();
    IPolicyDefinitionResourcesBuilder AddPresentation(PolicyPresentation                  presentation);
    IPolicyDefinitionResourcesBuilder AddPresentations(IEnumerable<PolicyPresentation>    presentations);
    IPolicyDefinitionResourcesBuilder RemovePresentation(PolicyPresentation               presentation);
    IPolicyDefinitionResourcesBuilder RemovePresentations(IEnumerable<PolicyPresentation> presentations);
    IPolicyDefinitionResourcesBuilder ClearPresentations();
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

        Revision = revision;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithSchemaVersion(string schemaVersion)
    {
        ArgumentNullException.ThrowIfNull(schemaVersion, nameof(schemaVersion));
        EnsureNotBuilt(nameof(WithSchemaVersion));

        SchemaVersion = schemaVersion;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName = displayName;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder WithDescription(string description)
    {
        ArgumentNullException.ThrowIfNull(description, nameof(description));
        EnsureNotBuilt(nameof(WithDescription));

        Description = description;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(AddAnnotation));
        _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(AddAnnotations));

        foreach (var annotation in annotations)
            _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemoveAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(RemoveAnnotation));
        _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemoveAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(RemoveAnnotations));

        foreach (var annotation in annotations)
            _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder ClearAnnotations()
    {
        EnsureNotBuilt(nameof(ClearAnnotations));
        _annotations.Clear();

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddString(LocalizedString localizedString)
    {
        EnsureNotBuilt(nameof(AddString));
        _strings.Add(localizedString);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddStrings(IEnumerable<LocalizedString> localizedStrings)
    {
        ArgumentNullException.ThrowIfNull(localizedStrings, nameof(localizedStrings));
        EnsureNotBuilt(nameof(AddStrings));

        foreach (var localizedString in localizedStrings)
            _strings.Add(localizedString);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemoveString(LocalizedString localizedString)
    {
        EnsureNotBuilt(nameof(RemoveString));
        _strings.Remove(localizedString);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemoveStrings(IEnumerable<LocalizedString> localizedStrings)
    {
        ArgumentNullException.ThrowIfNull(localizedStrings, nameof(localizedStrings));
        EnsureNotBuilt(nameof(RemoveStrings));

        foreach (var localizedString in localizedStrings)
            _strings.Remove(localizedString);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder ClearStrings()
    {
        EnsureNotBuilt(nameof(ClearStrings));
        _strings.Clear();
        _stringsView = null;

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddPresentation(PolicyPresentation presentation)
    {
        ArgumentNullException.ThrowIfNull(presentation, nameof(presentation));
        EnsureNotBuilt(nameof(AddPresentation));

        _presentations.Add(presentation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder AddPresentations(IEnumerable<PolicyPresentation> presentations)
    {
        ArgumentNullException.ThrowIfNull(presentations, nameof(presentations));
        EnsureNotBuilt(nameof(AddPresentations));

        foreach (var presentation in presentations)
        {
            if (presentation == null)
                throw new ArgumentNullException(nameof(presentations), "Presentations collection cannot contain null values.");

            _presentations.Add(presentation);
        }

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemovePresentation(PolicyPresentation presentation)
    {
        ArgumentNullException.ThrowIfNull(presentation, nameof(presentation));
        EnsureNotBuilt(nameof(RemovePresentation));

        _presentations.Remove(presentation);

        return this;
    }

    public IPolicyDefinitionResourcesBuilder RemovePresentations(IEnumerable<PolicyPresentation> presentations)
    {
        ArgumentNullException.ThrowIfNull(presentations, nameof(presentations));
        EnsureNotBuilt(nameof(RemovePresentations));

        foreach (var presentation in presentations)
        {
            if (presentation == null)
                throw new ArgumentNullException(nameof(presentations), "Presentations collection cannot contain null values.");

            _presentations.Remove(presentation);
        }

        return this;
    }

    public IPolicyDefinitionResourcesBuilder ClearPresentations()
    {
        EnsureNotBuilt(nameof(ClearPresentations));
        _presentations.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyDefinitionResourcesBuilder, PolicyDefinitionResources>(
            () => (nameof(Revision), Revision is not null),
            () => (nameof(SchemaVersion), SchemaVersion is not null),
            () => (nameof(DisplayName), DisplayName is not null),
            () => (nameof(Description), Description is not null)
        );

    protected override PolicyDefinitionResources BuildCore() =>
        new()
        {
            Revision      = Revision!,
            SchemaVersion = SchemaVersion!,
            DisplayName   = DisplayName!,
            Description   = Description!,
            Annotations   = Annotations,
            Strings       = Strings,
            Presentations = Presentations
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
    }
}