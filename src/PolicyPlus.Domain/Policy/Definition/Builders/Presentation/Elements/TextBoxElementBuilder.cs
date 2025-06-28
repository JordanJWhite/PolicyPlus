using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for TextBoxElement.
/// </summary>
public interface ITextBoxElementBuilder : IBuilder<TextBoxElement>
{
    ITextBoxElementBuilder WithLabel(string         label);
    ITextBoxElementBuilder WithDefaultValue(string? defaultValue);
    ITextBoxElementBuilder WithRefId(string         refId);
}

/// <summary>
///     Builder for TextBoxElement.
/// </summary>
public class TextBoxElementBuilder : BuilderBase<TextBoxElementBuilder, TextBoxElement>, ITextBoxElementBuilder
{
    public string? Label        { get; private set; }
    public string? DefaultValue { get; private set; }
    public string? RefId        { get; private set; }

    public ITextBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public ITextBoxElementBuilder WithDefaultValue(string? defaultValue)
    {
        EnsureNotBuilt(nameof(WithDefaultValue));
        DefaultValue = defaultValue;

        return this;
    }

    public ITextBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<TextBoxElementBuilder, TextBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override TextBoxElement BuildCore() =>
        new()
        {
            Label        = Label!,
            DefaultValue = DefaultValue,
            RefId        = RefId!
        };

    protected override void ResetCore()
    {
        Label        = null;
        DefaultValue = null;
        RefId        = null;
    }
}