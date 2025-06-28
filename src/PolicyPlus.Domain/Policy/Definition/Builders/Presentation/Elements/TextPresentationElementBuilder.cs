using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for TextPresentationElement.
/// </summary>
public interface ITextPresentationElementBuilder : IBuilder<TextPresentationElement>
{
    ITextPresentationElementBuilder WithText(string  text);
    ITextPresentationElementBuilder WithRefId(string refId);
}

/// <summary>
///     Builder for TextPresentationElement.
/// </summary>
public class TextPresentationElementBuilder
    : BuilderBase<TextPresentationElementBuilder, TextPresentationElement>, ITextPresentationElementBuilder
{
    public string? Text  { get; private set; }
    public string? RefId { get; private set; }

    public ITextPresentationElementBuilder WithText(string text)
    {
        ArgumentNullException.ThrowIfNull(text, nameof(text));
        EnsureNotBuilt(nameof(WithText));

        Text = text;

        return this;
    }

    public ITextPresentationElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<TextPresentationElementBuilder, TextPresentationElement>(
            () => (nameof(Text), Text is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override TextPresentationElement BuildCore() =>
        new()
        {
            Text  = Text!,
            RefId = RefId!
        };

    protected override void ResetCore()
    {
        Text  = null;
        RefId = null;
    }
}