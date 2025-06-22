using PolicyPlus.Domain.Policy.Definition.Core;
using PolicyPlus.Domain.Policy.Definition.Elements;
using PolicyPlus.Domain.Policy.Definition.References;
using PolicyPlus.Domain.Policy.Definition.SupportedOn;
using PolicyPlus.Domain.Policy.Definition.Values;

astructure.Files.GroupPolicy.Builders;

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