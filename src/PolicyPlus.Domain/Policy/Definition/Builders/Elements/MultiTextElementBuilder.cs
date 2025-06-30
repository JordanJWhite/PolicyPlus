using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for MultiTextElement.
/// </summary>
public interface IMultiTextElementBuilder : IBuilder<MultiTextElement>
{
    IMultiTextElementBuilder WithId(string               id);
    IMultiTextElementBuilder WithClientExtension(string? clientExtension);
    IMultiTextElementBuilder WithKey(string?             key);
    IMultiTextElementBuilder WithValueName(string?       valueName);
    IMultiTextElementBuilder WithRequired(bool           required);
    IMultiTextElementBuilder WithMaxLength(uint          maxLength);
    IMultiTextElementBuilder WithMaxStrings(uint         maxStrings);
    IMultiTextElementBuilder WithSoft(bool               soft);
}

/// <summary>
///     Builder for MultiTextElement.
/// </summary>
public class MultiTextElementBuilder : BuilderBase<MultiTextElementBuilder, MultiTextElement>, IMultiTextElementBuilder
{
    public string? Id              { get; private set; }
    public string? ClientExtension { get; private set; }
    public string? Key             { get; private set; }
    public string? ValueName       { get; private set; }
    public bool    Required        { get; private set; }
    public uint    MaxLength       { get; private set; } = 1023;
    public uint    MaxStrings      { get; private set; }
    public bool    Soft            { get; private set; }

    public IMultiTextElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));
        Id = id;

        return this;
    }

    public IMultiTextElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public IMultiTextElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IMultiTextElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public IMultiTextElementBuilder WithRequired(bool required)
    {
        EnsureNotBuilt(nameof(WithRequired));
        Required = required;

        return this;
    }

    public IMultiTextElementBuilder WithMaxLength(uint maxLength)
    {
        EnsureNotBuilt(nameof(WithMaxLength));
        MaxLength = maxLength;

        return this;
    }

    public IMultiTextElementBuilder WithMaxStrings(uint maxStrings)
    {
        EnsureNotBuilt(nameof(WithMaxStrings));
        MaxStrings = maxStrings;

        return this;
    }

    public IMultiTextElementBuilder WithSoft(bool soft)
    {
        EnsureNotBuilt(nameof(WithSoft));
        Soft = soft;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<MultiTextElementBuilder, MultiTextElement>(() => (nameof(Id), Id is not null));

    protected override MultiTextElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValueName       = ValueName,
            Required        = Required,
            MaxLength       = MaxLength,
            MaxStrings      = MaxStrings,
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
        MaxStrings      = 0;
        Soft            = false;
    }
}