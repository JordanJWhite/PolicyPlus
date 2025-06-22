using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for TextElement.
/// </summary>
public interface ITextElementBuilder : IBuilder<TextElement>
{
    ITextElementBuilder WithId(string               id);
    ITextElementBuilder WithClientExtension(string? clientExtension);
    ITextElementBuilder WithKey(string?             key);
    ITextElementBuilder WithValueName(string?       valueName);
    ITextElementBuilder WithRequired(bool           required);
    ITextElementBuilder WithMaxLength(uint          maxLength);
    ITextElementBuilder WithExpandable(bool         expandable);
    ITextElementBuilder WithSoft(bool               soft);
}

/// <summary>
///     Builder for TextElement.
/// </summary>
public class TextElementBuilder : BuilderBase<TextElementBuilder, TextElement>, ITextElementBuilder
{
    public string? Id              { get; private set; }
    public string? ClientExtension { get; private set; }
    public string? ValueName       { get; private set; }
    public string? Key             { get; private set; }
    public bool    Required        { get; private set; }
    public uint    MaxLength       { get; private set; } = 1023;
    public bool    Expandable      { get; private set; }
    public bool    Soft            { get; private set; }

    public ITextElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));
        Id = id;

        return this;
    }

    public ITextElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public ITextElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public ITextElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public ITextElementBuilder WithRequired(bool required)
    {
        EnsureNotBuilt(nameof(WithRequired));
        Required = required;

        return this;
    }

    public ITextElementBuilder WithMaxLength(uint maxLength)
    {
        EnsureNotBuilt(nameof(WithMaxLength));
        MaxLength = maxLength;

        return this;
    }

    public ITextElementBuilder WithExpandable(bool expandable)
    {
        EnsureNotBuilt(nameof(WithExpandable));
        Expandable = expandable;

        return this;
    }

    public ITextElementBuilder WithSoft(bool soft)
    {
        EnsureNotBuilt(nameof(WithSoft));
        Soft = soft;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<TextElementBuilder, TextElement>(() => (nameof(Id), Id is not null));

    protected override TextElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValueName       = ValueName,
            Required        = Required,
            MaxLength       = MaxLength,
            Expandable      = Expandable,
            Soft            = Soft
        };

    protected override void ResetCore()
    {
        Id              = null;
        ClientExtension = null;
        Key             = null;
        ValueName       = null;
        Required        = false;
        MaxLength       = 1023;
        Expandable      = false;
        Soft            = false;
    }
}