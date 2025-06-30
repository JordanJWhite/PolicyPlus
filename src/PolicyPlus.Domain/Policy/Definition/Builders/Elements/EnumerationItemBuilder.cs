using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for EnumerationItem.
/// </summary>
public interface IEnumerationItemBuilder : IBuilder<EnumerationItem>
{
    /// <summary>
    ///     Sets the display name reference for the enumeration item.
    /// </summary>
    /// <param name="displayName">The display name reference (e.g., "$(string.optionName)").</param>
    /// <returns>The builder instance for method chaining.</returns>
    IEnumerationItemBuilder WithDisplayName(string displayName);

    /// <summary>
    ///     Sets the value for the enumeration item.
    /// </summary>
    /// <param name="value">The value to set when this item is selected.</param>
    /// <returns>The builder instance for method chaining.</returns>
    IEnumerationItemBuilder WithValue(IValue value);

    /// <summary>
    ///     Sets the optional value list for the enumeration item.
    /// </summary>
    /// <param name="valueList">The value list containing additional values to set when this item is selected.</param>
    /// <returns>The builder instance for method chaining.</returns>
    IEnumerationItemBuilder WithValueList(ValueList? valueList);
}

/// <summary>
///     Builder for EnumerationItem.
/// </summary>
public class EnumerationItemBuilder : BuilderBase<EnumerationItemBuilder, EnumerationItem>, IEnumerationItemBuilder
{
    public string?    DisplayName { get; private set; }
    public IValue?    Value       { get; private set; }
    public ValueList? ValueList   { get; private set; }

    public IEnumerationItemBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));
        DisplayName = displayName;

        return this;
    }

    public IEnumerationItemBuilder WithValue(IValue value)
    {
        ArgumentNullException.ThrowIfNull(value, nameof(value));
        EnsureNotBuilt(nameof(WithValue));

        Value = value;

        return this;
    }

    public IEnumerationItemBuilder WithValueList(ValueList? valueList)
    {
        EnsureNotBuilt(nameof(WithValueList));
        ValueList = valueList;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<EnumerationItemBuilder, EnumerationItem>(
            () => (nameof(DisplayName), DisplayName is not null),
            () => (nameof(Value), Value is not null)
        );

    protected override EnumerationItem BuildCore() =>
        new()
        {
            DisplayName = DisplayName!,
            Value       = Value!,
            ValueList   = ValueList
        };

    protected override void ResetCore()
    {
        DisplayName = null;
        Value       = null;
        ValueList   = null;
    }
}