using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation;

/// <summary>
///     Builder interface for PolicyPresentation.
/// </summary>
public interface IPolicyPresentationBuilder : IBuilder<PolicyPresentation>
{
    IPolicyPresentationBuilder WithId(string                   id);
    IPolicyPresentationBuilder AddElement(IPresentationElement element);
}

/// <summary>
///     Builder for PolicyPresentation.
/// </summary>
public class PolicyPresentationBuilder : BuilderBase<PolicyPresentationBuilder, PolicyPresentation>, IPolicyPresentationBuilder
{
    private List<IPresentationElement>           _elements = [];
    private IReadOnlyList<IPresentationElement>? _elementsView;

    public string?                             Id       { get; private set; }
    public IReadOnlyList<IPresentationElement> Elements => _elementsView ??= _elements.AsReadOnly();

    public IPolicyPresentationBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));

        Id = id;

        return this;
    }

    public IPolicyPresentationBuilder AddElement(IPresentationElement element)
    {
        ArgumentNullException.ThrowIfNull(element, nameof(element));
        EnsureNotBuilt(nameof(AddElement));

        _elements.Add(element);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<PolicyPresentationBuilder, PolicyPresentation>(() => (nameof(Id), Id is not null));

    protected override PolicyPresentation BuildCore() =>
        new()
        {
            Id       = Id!,
            Elements = Elements
        };

    protected override void ResetCore()
    {
        Id            = null;
        _elements     = [];
        _elementsView = null;
    }
}