using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Core;
using PolicyPlus.Domain.Policy.Definition.Elements;
using PolicyPlus.Domain.Policy.Definition.References;
using PolicyPlus.Domain.Policy.Definition.SupportedOn;
using PolicyPlus.Domain.Policy.Definition.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Core;

/// <summary>
///     Builder interface for PolicyDefinition.
/// </summary>
public interface IPolicyDefinitionBuilder : IBuilder<PolicyDefinition>
{
    IPolicyDefinitionBuilder WithName(string                       name);
    IPolicyDefinitionBuilder WithClass(PolicyClass                 policyClass);
    IPolicyDefinitionBuilder WithDisplayName(string                displayName);
    IPolicyDefinitionBuilder WithExplainText(string?               explainText);
    IPolicyDefinitionBuilder WithKey(string                        key);
    IPolicyDefinitionBuilder WithValueName(string?                 valueName);
    IPolicyDefinitionBuilder WithParentCategory(CategoryReference? parentCategory);
    IPolicyDefinitionBuilder WithSupportedOn(SupportedOnReference  supportedOn);
    IPolicyDefinitionBuilder WithPresentation(string?              presentation);
    IPolicyDefinitionBuilder WithEnabledValue(IValue?              enabledValue);
    IPolicyDefinitionBuilder WithDisabledValue(IValue?             disabledValue);
    IPolicyDefinitionBuilder WithEnabledList(ValueList?            enabledList);
    IPolicyDefinitionBuilder WithDisabledList(ValueList?           disabledList);
    IPolicyDefinitionBuilder WithKeywords(string?                  keywords);
    IPolicyDefinitionBuilder AddAnnotation(Annotation              annotation);
    IPolicyDefinitionBuilder AddSeeAlso(string                     seeAlso);
    IPolicyDefinitionBuilder AddElement(PolicyElementBase          element);
}

/// <summary>
///     Builder for PolicyDefinition.
/// </summary>
public class PolicyDefinitionBuilder : BuilderBase<PolicyDefinitionBuilder, PolicyDefinition>, IPolicyDefinitionBuilder
{
    private readonly List<Annotation> _annotations = new();
    private readonly List<PolicyElementBase> _elements = new();
    private readonly List<string> _seeAlso = new();
    
    private bool _nameSet;
    private bool _classSet;
    private bool _displayNameSet;
    private bool _keySet;
    private bool _supportedOnSet;

    public string? Name { get; private set; }
    public PolicyClass? Class { get; private set; }
    public string? DisplayName { get; private set; }
    public string? ExplainText { get; private set; }
    public string? Key { get; private set; }
    public string? ValueName { get; private set; }
    public CategoryReference? ParentCategory { get; private set; }
    public SupportedOnReference? SupportedOn { get; private set; }
    public string? Presentation { get; private set; }
    public IValue? EnabledValue { get; private set; }
    public IValue? DisabledValue { get; private set; }
    public ValueList? EnabledList { get; private set; }
    public ValueList? DisabledList { get; private set; }
    public string? Keywords { get; private set; }
    public IReadOnlyList<Annotation> Annotations => _annotations.AsReadOnly();
    public IReadOnlyList<string> SeeAlso => _seeAlso.AsReadOnly();
    public IReadOnlyList<PolicyElementBase> Elements => _elements.AsReadOnly();

    public IPolicyDefinitionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));
        
        Name = name;
        _nameSet = true;

        return this;
    }
}