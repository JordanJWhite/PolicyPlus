using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for CheckBoxElement.
/// </summary>
public interface ICheckBoxElementBuilder : IBuilder<CheckBoxElement>
{
    ICheckBoxElementBuilder WithLabel(string        label);
    ICheckBoxElementBuilder WithDefaultChecked(bool defaultChecked);
    ICheckBoxElementBuilder WithRefId(string        refId);
}

/// <summary>
///     Builder for CheckBoxElement.
/// </summary>
public class CheckBoxElementBuilder : BuilderBase<CheckBoxElementBuilder, CheckBoxElement>, ICheckBoxElementBuilder
{
    public string? Label          { get; private set; }
    public bool    DefaultChecked { get; private set; }
    public string? RefId          { get; private set; }

    public ICheckBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public ICheckBoxElementBuilder WithDefaultChecked(bool defaultChecked)
    {
        EnsureNotBuilt(nameof(WithDefaultChecked));
        DefaultChecked = defaultChecked;

        return this;
    }

    public ICheckBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<CheckBoxElementBuilder, CheckBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override CheckBoxElement BuildCore() =>
        new()
        {
            Label          = Label!,
            DefaultChecked = DefaultChecked,
            RefId          = RefId!
        };

    protected override void ResetCore()
    {
        Label          = null;
        DefaultChecked = false;
        RefId          = null;
    }
}