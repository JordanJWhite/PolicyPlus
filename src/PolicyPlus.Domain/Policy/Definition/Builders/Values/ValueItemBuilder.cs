using PolicyPlus.Domain.Policy.Definition.Model.Values;
using PolicyPlus.Domain.Utility.Builders.Base;
using PolicyPlus.Domain.Utility.Builders.Helpers;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for ValueItem.
/// </summary>
public interface IValueItemBuilder : IBuilder<ValueItem>
{
    IValueItemBuilder WithValue(IValue     value);
    IValueItemBuilder WithKey(string?      key);
    IValueItemBuilder WithValueName(string valueName);
}

/// <summary>
///     Builder for ValueItem.
/// </summary>
public class ValueItemBuilder : BuilderBase<ValueItemBuilder, ValueItem>, IValueItemBuilder
{
    public IValue? Value     { get; private set; }
    public string? Key       { get; private set; }
    public string? ValueName { get; private set; }

    public IValueItemBuilder WithValue(IValue value)
    {
        ArgumentNullException.ThrowIfNull(value, nameof(value));
        EnsureNotBuilt(nameof(WithValue));
        Value = value;

        return this;
    }

    public IValueItemBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IValueItemBuilder WithValueName(string valueName)
    {
        ArgumentNullException.ThrowIfNull(valueName, nameof(valueName));
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<ValueItemBuilder, ValueItem>(
            () => (nameof(ValueName), ValueName is not null),
            () => (nameof(Value), Value is not null)
        );

    protected override ValueItem BuildCore() =>
        new()
        {
            Value     = Value!,
            Key       = Key,
            ValueName = ValueName!
        };

    protected override void ResetCore()
    {
        Value     = null;
        Key       = null;
        ValueName = null;
    }
}