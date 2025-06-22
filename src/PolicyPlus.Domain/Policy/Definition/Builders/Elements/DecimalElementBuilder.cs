using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for DecimalElement.
/// </summary>
public interface IDecimalElementBuilder : IBuilder<DecimalElement>
{
    IDecimalElementBuilder WithId(string               id);
    IDecimalElementBuilder WithClientExtension(string? clientExtension);
    IDecimalElementBuilder WithKey(string?             key);
    IDecimalElementBuilder WithValueName(string?       valueName);
    IDecimalElementBuilder WithRequired(bool           required);
    IDecimalElementBuilder WithMinValue(uint           minValue);
    IDecimalElementBuilder WithMaxValue(uint           maxValue);
    IDecimalElementBuilder WithStoreAsText(bool        storeAsText);
    IDecimalElementBuilder WithSoft(bool               soft);
}

/// <summary>
///     Builder for DecimalElement.
/// </summary>
public class DecimalElementBuilder : BuilderBase<DecimalElementBuilder, DecimalElement>, IDecimalElementBuilder
{
    public string? ClientExtension { get; private set; }
    public string? Id              { get; private set; }
    public string? Key             { get; private set; }
    public uint    MaxValue        { get; private set; } = 9999;
    public uint    MinValue        { get; private set; }
    public bool    Required        { get; private set; }
    public bool    Soft            { get; private set; }
    public bool    StoreAsText     { get; private set; }
    public string? ValueName       { get; private set; }

    public IDecimalElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));

        Id = id;

        return this;
    }

    public IDecimalElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public IDecimalElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IDecimalElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public IDecimalElementBuilder WithRequired(bool required)
    {
        EnsureNotBuilt(nameof(WithRequired));
        Required = required;

        return this;
    }

    public IDecimalElementBuilder WithMinValue(uint minValue)
    {
        EnsureNotBuilt(nameof(WithMinValue));
        MinValue = minValue;

        return this;
    }

    public IDecimalElementBuilder WithMaxValue(uint maxValue)
    {
        EnsureNotBuilt(nameof(WithMaxValue));
        MaxValue = maxValue;

        return this;
    }

    public IDecimalElementBuilder WithStoreAsText(bool storeAsText)
    {
        EnsureNotBuilt(nameof(WithStoreAsText));
        StoreAsText = storeAsText;

        return this;
    }

    public IDecimalElementBuilder WithSoft(bool soft)
    {
        EnsureNotBuilt(nameof(WithSoft));
        Soft = soft;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<DecimalElementBuilder, DecimalElement>(() => (nameof(Id), Id is not null));

    protected override DecimalElement BuildCore() =>
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