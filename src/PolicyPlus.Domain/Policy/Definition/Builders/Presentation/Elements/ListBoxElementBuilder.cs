using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for ListBoxElement.
/// </summary>
public interface IListBoxElementBuilder : IBuilder<ListBoxElement>
{
    IListBoxElementBuilder WithLabel(string label);
    IListBoxElementBuilder WithRefId(string refId);
}

/// <summary>
///     Builder for ListBoxElement.
/// </summary>
public class ListBoxElementBuilder : BuilderBase<ListBoxElementBuilder, ListBoxElement>, IListBoxElementBuilder
{
    public string? Label { get; private set; }
    public string? RefId { get; private set; }

    public IListBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public IListBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<ListBoxElementBuilder, ListBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override ListBoxElement BuildCore() =>
        new()
        {
            Label = Label!,
            RefId = RefId!
        };

    protected override void ResetCore()
    {
        Label = null;
        RefId = null;
    }
}