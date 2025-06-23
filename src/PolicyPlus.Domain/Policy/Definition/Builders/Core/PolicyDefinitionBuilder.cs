using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;
using PolicyPlus.Domain.Policy.Definition.Model.References;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Core;

/// <summary>
///     Builder interface for PolicyDefinition.
/// </summary>
public interface IPolicyDefinitionBuilder : IBuilder<PolicyDefinition>
{
    IPolicyDefinitionBuilder WithName(string                           name);
    IPolicyDefinitionBuilder WithClass(PolicyClass                     policyClass);
    IPolicyDefinitionBuilder WithDisplayName(string                    displayName);
    IPolicyDefinitionBuilder WithExplainText(string?                   explainText);
    IPolicyDefinitionBuilder WithKey(string                            key);
    IPolicyDefinitionBuilder WithValueName(string?                     valueName);
    IPolicyDefinitionBuilder WithParentCategory(CategoryReference?     parentCategory);
    IPolicyDefinitionBuilder WithSupportedOn(SupportedOnReference      supportedOn);
    IPolicyDefinitionBuilder WithPresentation(string?                  presentation);
    IPolicyDefinitionBuilder WithEnabledValue(IValue?                  enabledValue);
    IPolicyDefinitionBuilder WithDisabledValue(IValue?                 disabledValue);
    IPolicyDefinitionBuilder WithEnabledList(ValueList?                enabledList);
    IPolicyDefinitionBuilder WithDisabledList(ValueList?               disabledList);
    IPolicyDefinitionBuilder WithKeywords(string?                      keywords);
    IPolicyDefinitionBuilder AddAnnotation(Annotation                  annotation);
    IPolicyDefinitionBuilder AddAnnotations(IEnumerable<Annotation>    annotations);
    IPolicyDefinitionBuilder RemoveAnnotation(Annotation               annotation);
    IPolicyDefinitionBuilder RemoveAnnotations(IEnumerable<Annotation> annotations);
    IPolicyDefinitionBuilder ClearAnnotations();
    IPolicyDefinitionBuilder AddSeeAlso(string                  seeAlso);
    IPolicyDefinitionBuilder AddSeeAlsos(IEnumerable<string>    seeAlsos);
    IPolicyDefinitionBuilder RemoveSeeAlso(string               seeAlso);
    IPolicyDefinitionBuilder RemoveSeeAlsos(IEnumerable<string> seeAlsos);
    IPolicyDefinitionBuilder ClearSeeAlsos();
    IPolicyDefinitionBuilder AddElement(PolicyElementBase                  element);
    IPolicyDefinitionBuilder AddElements(IEnumerable<PolicyElementBase>    elements);
    IPolicyDefinitionBuilder RemoveElement(PolicyElementBase               element);
    IPolicyDefinitionBuilder RemoveElements(IEnumerable<PolicyElementBase> elements);
    IPolicyDefinitionBuilder ClearElements();
}

/// <summary>
///     Builder for PolicyDefinition.
/// </summary>
public class PolicyDefinitionBuilder : BuilderBase<PolicyDefinitionBuilder, PolicyDefinition>, IPolicyDefinitionBuilder
{
    // Collection fields for annotations, elements, and see-also links.
    private List<Annotation>        _annotations = [];
    private List<PolicyElementBase> _elements    = [];
    private List<string>            _seeAlso     = [];

    // Views for the collections to provide read-only access.
    private IReadOnlyList<Annotation>?        _annotationsView;
    private IReadOnlyList<PolicyElementBase>? _elementsView;
    private IReadOnlyList<string>?            _seeAlsoView;

    // Flags to track whether required properties have been set.
    private bool _nameSet;
    private bool _classSet;
    private bool _displayNameSet;
    private bool _keySet;
    private bool _supportedOnSet;

    public string?                          Name           { get; private set; }
    public PolicyClass?                     Class          { get; private set; }
    public string?                          DisplayName    { get; private set; }
    public string?                          ExplainText    { get; private set; }
    public string?                          Key            { get; private set; }
    public string?                          ValueName      { get; private set; }
    public CategoryReference?               ParentCategory { get; private set; }
    public SupportedOnReference?            SupportedOn    { get; private set; }
    public string?                          Presentation   { get; private set; }
    public IValue?                          EnabledValue   { get; private set; }
    public IValue?                          DisabledValue  { get; private set; }
    public ValueList?                       EnabledList    { get; private set; }
    public ValueList?                       DisabledList   { get; private set; }
    public string?                          Keywords       { get; private set; }
    public IReadOnlyList<Annotation>        Annotations    => _annotationsView ??= _annotations.AsReadOnly();
    public IReadOnlyList<string>            SeeAlso        => _seeAlsoView ??= _seeAlso.AsReadOnly();
    public IReadOnlyList<PolicyElementBase> Elements       => _elementsView ??= _elements.AsReadOnly();

    public IPolicyDefinitionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));

        Name     = name;
        _nameSet = true;

        return this;
    }

    public IPolicyDefinitionBuilder WithClass(PolicyClass policyClass)
    {
        EnsureNotBuilt(nameof(WithClass));

        Class     = policyClass;
        _classSet = true;

        return this;
    }

    public IPolicyDefinitionBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName     = displayName;
        _displayNameSet = true;

        return this;
    }

    public IPolicyDefinitionBuilder WithExplainText(string? explainText)
    {
        EnsureNotBuilt(nameof(WithExplainText));
        ExplainText = explainText;

        return this;
    }

    public IPolicyDefinitionBuilder WithKey(string key)
    {
        ArgumentNullException.ThrowIfNull(key, nameof(key));
        EnsureNotBuilt(nameof(WithKey));

        Key     = key;
        _keySet = true;

        return this;
    }

    public IPolicyDefinitionBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public IPolicyDefinitionBuilder WithParentCategory(CategoryReference? parentCategory)
    {
        EnsureNotBuilt(nameof(WithParentCategory));
        ParentCategory = parentCategory;

        return this;
    }

    public IPolicyDefinitionBuilder WithSupportedOn(SupportedOnReference supportedOn)
    {
        EnsureNotBuilt(nameof(WithSupportedOn));

        SupportedOn     = supportedOn;
        _supportedOnSet = true;

        return this;
    }

    public IPolicyDefinitionBuilder WithPresentation(string? presentation)
    {
        EnsureNotBuilt(nameof(WithPresentation));
        Presentation = presentation;

        return this;
    }

    public IPolicyDefinitionBuilder WithEnabledValue(IValue? enabledValue)
    {
        EnsureNotBuilt(nameof(WithEnabledValue));
        EnabledValue = enabledValue;

        return this;
    }

    public IPolicyDefinitionBuilder WithDisabledValue(IValue? disabledValue)
    {
        EnsureNotBuilt(nameof(WithDisabledValue));
        DisabledValue = disabledValue;

        return this;
    }

    public IPolicyDefinitionBuilder WithEnabledList(ValueList? enabledList)
    {
        EnsureNotBuilt(nameof(WithEnabledList));
        EnabledList = enabledList;

        return this;
    }

    public IPolicyDefinitionBuilder WithDisabledList(ValueList? disabledList)
    {
        EnsureNotBuilt(nameof(WithDisabledList));
        DisabledList = disabledList;

        return this;
    }

    public IPolicyDefinitionBuilder WithKeywords(string? keywords)
    {
        EnsureNotBuilt(nameof(WithKeywords));
        Keywords = keywords;

        return this;
    }

    public IPolicyDefinitionBuilder AddAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(AddAnnotation));
        _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionBuilder AddAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(AddAnnotations));

        foreach (var annotation in annotations)
            _annotations.Add(annotation);

        return this;
    }

    public IPolicyDefinitionBuilder RemoveAnnotation(Annotation annotation)
    {
        EnsureNotBuilt(nameof(RemoveAnnotation));
        _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionBuilder RemoveAnnotations(IEnumerable<Annotation> annotations)
    {
        ArgumentNullException.ThrowIfNull(annotations, nameof(annotations));
        EnsureNotBuilt(nameof(RemoveAnnotations));

        foreach (var annotation in annotations)
            _annotations.Remove(annotation);

        return this;
    }

    public IPolicyDefinitionBuilder ClearAnnotations()
    {
        EnsureNotBuilt(nameof(ClearAnnotations));
        _annotations.Clear();

        return this;
    }

    public IPolicyDefinitionBuilder AddSeeAlso(string seeAlso)
    {
        ArgumentNullException.ThrowIfNull(seeAlso, nameof(seeAlso));
        EnsureNotBuilt(nameof(AddSeeAlso));

        _seeAlso.Add(seeAlso);

        return this;
    }

    public IPolicyDefinitionBuilder AddSeeAlsos(IEnumerable<string> seeAlsos)
    {
        ArgumentNullException.ThrowIfNull(seeAlsos, nameof(seeAlsos));
        EnsureNotBuilt(nameof(AddSeeAlsos));

        foreach (var seeAlso in seeAlsos)
        {
            if (seeAlso == null)
                throw new ArgumentNullException(nameof(seeAlsos), "SeeAlsos collection cannot contain null values.");

            _seeAlso.Add(seeAlso);
        }

        return this;
    }

    public IPolicyDefinitionBuilder RemoveSeeAlso(string seeAlso)
    {
        ArgumentNullException.ThrowIfNull(seeAlso, nameof(seeAlso));
        EnsureNotBuilt(nameof(RemoveSeeAlso));

        _seeAlso.Remove(seeAlso);

        return this;
    }

    public IPolicyDefinitionBuilder RemoveSeeAlsos(IEnumerable<string> seeAlsos)
    {
        ArgumentNullException.ThrowIfNull(seeAlsos, nameof(seeAlsos));
        EnsureNotBuilt(nameof(RemoveSeeAlsos));

        foreach (var seeAlso in seeAlsos)
        {
            if (seeAlso == null)
                throw new ArgumentNullException(nameof(seeAlsos), "SeeAlsos collection cannot contain null values.");

            _seeAlso.Remove(seeAlso);
        }

        return this;
    }

    public IPolicyDefinitionBuilder ClearSeeAlsos()
    {
        EnsureNotBuilt(nameof(ClearSeeAlsos));
        _seeAlso.Clear();

        return this;
    }

    public IPolicyDefinitionBuilder AddElement(PolicyElementBase element)
    {
        ArgumentNullException.ThrowIfNull(element, nameof(element));
        EnsureNotBuilt(nameof(AddElement));

        _elements.Add(element);

        return this;
    }

    public IPolicyDefinitionBuilder AddElements(IEnumerable<PolicyElementBase> elements)
    {
        ArgumentNullException.ThrowIfNull(elements, nameof(elements));
        EnsureNotBuilt(nameof(AddElements));

        foreach (var element in elements)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(elements), "Elements collection cannot contain null values.");

            _elements.Add(element);
        }

        return this;
    }

    public IPolicyDefinitionBuilder RemoveElement(PolicyElementBase element)
    {
        ArgumentNullException.ThrowIfNull(element, nameof(element));
        EnsureNotBuilt(nameof(RemoveElement));

        _elements.Remove(element);

        return this;
    }

    public IPolicyDefinitionBuilder RemoveElements(IEnumerable<PolicyElementBase> elements)
    {
        ArgumentNullException.ThrowIfNull(elements, nameof(elements));
        EnsureNotBuilt(nameof(RemoveElements));

        foreach (var element in elements)
        {
            if (element == null)
                throw new ArgumentNullException(nameof(elements), "Elements collection cannot contain null values.");

            _elements.Remove(element);
        }

        return this;
    }

    public IPolicyDefinitionBuilder ClearElements()
    {
        EnsureNotBuilt(nameof(ClearElements));
        _elements.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<PolicyDefinitionBuilder, PolicyDefinition>(
            () => (nameof(Name), _nameSet),
            () => (nameof(Class), _classSet),
            () => (nameof(DisplayName), _displayNameSet),
            () => (nameof(Key), _keySet),
            () => (nameof(SupportedOn), _supportedOnSet)
        );

    protected override PolicyDefinition BuildCore() =>
        new()
        {
            Name           = Name!,
            Class          = Class!.Value,
            DisplayName    = DisplayName!,
            ExplainText    = ExplainText,
            Key            = Key!,
            ValueName      = ValueName,
            ParentCategory = ParentCategory,
            SupportedOn    = SupportedOn!.Value,
            Presentation   = Presentation,
            EnabledValue   = EnabledValue,
            DisabledValue  = DisabledValue,
            EnabledList    = EnabledList,
            DisabledList   = DisabledList,
            Keywords       = Keywords,
            Annotations    = _annotations.AsReadOnly(),
            SeeAlso        = _seeAlso.AsReadOnly(),
            Elements       = _elements.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Name             = null;
        Class            = null;
        DisplayName      = null;
        ExplainText      = null;
        Key              = null;
        ValueName        = null;
        ParentCategory   = null;
        SupportedOn      = null;
        Presentation     = null;
        EnabledValue     = null;
        DisabledValue    = null;
        EnabledList      = null;
        DisabledList     = null;
        Keywords         = null;
        _annotations     = [];
        _elements        = [];
        _seeAlso         = [];
        _annotationsView = null;
        _elementsView    = null;
        _seeAlsoView     = null;
        _nameSet         = false;
        _classSet        = false;
        _displayNameSet  = false;
        _keySet          = false;
        _supportedOnSet  = false;
    }
}