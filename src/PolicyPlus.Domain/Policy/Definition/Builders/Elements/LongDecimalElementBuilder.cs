using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for LongDecimalElement.
/// </summary>
public interface ILongDecimalElementBuilder : IBuilder<LongDecimalElement>
{
    ILongDecimalElementBuilder WithId(string               id);
    ILongDecimalElementBuilder WithClientExtension(string? clientExtension);
    ILongDecimalElementBuilder WithKey(string?             key);
    ILongDecimalElementBuilder WithValueName(string?       valueName);
    ILongDecimalElementBuilder WithRequired(bool           required);
    ILongDecimalElementBuilder WithMinValue(ulong          minValue);
    ILongDecimalElementBuilder WithMaxValue(ulong          maxValue);
    ILongDecimalElementBuilder WithStoreAsText(bool        storeAsText);
    ILongDecimalElementBuilder WithSoft(bool               soft);
}

/// <summary>
///     Builder for LongDecimalElement.
/// </summary>
public class LongDecimalElementBuilder : BuilderBase<LongDecimalElementBuilder, LongDecimalElement>, ILongDecimalElementBuilder
{
    public string? Id              { get; private set; }
    public string? ClientExtension { get; private set; }
    public string? Key             { get; private set; }
    public string? ValueName       { get; private set; }
    public bool    Required        { get; private set; }
    public ulong   MinValue        { get; private set; }
    public ulong   MaxValue        { get; private set; } = 9999;
    public bool    StoreAsText     { get; private set; }
    public bool    Soft            { get; private set; }

    public ILongDecimalElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));

        Id = id;

        return this;
    }

    public ILongDecimalElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public ILongDecimalElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public ILongDecimalElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public ILongDecimalElementBuilder WithRequired(bool required)
    {
        EnsureNotBuilt(nameof(WithRequired));
        Required = required;

        return this;
    }

    public ILongDecimalElementBuilder WithMinValue(ulong minValue)
    {
        EnsureNotBuilt(nameof(WithMinValue));
        MinValue = minValue;

        return this;
    }

    public ILongDecimalElementBuilder WithMaxValue(ulong maxValue)
    {
        EnsureNotBuilt(nameof(WithMaxValue));
        MaxValue = maxValue;

        return this;
    }

    public ILongDecimalElementBuilder WithStoreAsText(bool storeAsText)
    {
        EnsureNotBuilt(nameof(WithStoreAsText));
        StoreAsText = storeAsText;

        return this;
    }

    public ILongDecimalElementBuilder WithSoft(bool soft)
    {
        EnsureNotBuilt(nameof(WithSoft));
        Soft = soft;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<LongDecimalElementBuilder, LongDecimalElement>(() => (nameof(Id), Id is not null));

    protected override LongDecimalElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValueName       = ValueName,
            Required        = Required,
            MinValue        = MinValue,
            MaxValue        = MaxValue,
            StoreAsText     = StoreAsText,
            Soft            = Soft
        };

    protected override void ResetCore()
    {
        Id              = null;
        ClientExtension = null;
        Key             = null;
        ValueName       = null;
        Required        = false;
        MinValue        = 0;
        MaxValue        = 9999;
        StoreAsText     = false;
        Soft            = false;
    }
}