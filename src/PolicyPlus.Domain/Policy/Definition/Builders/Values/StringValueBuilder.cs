using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for StringValue.
/// </summary>
public interface IStringValueBuilder : IBuilder<StringValue>
{
    /// <summary>
    ///     Sets the string value.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    IStringValueBuilder WithValue(string value);
}

/// <summary>
///     Builder for StringValue.
/// </summary>
public class StringValueBuilder : BuilderBase<StringValueBuilder, StringValue>, IStringValueBuilder
{
    /// <summary>
    ///     Gets the string value.
    /// </summary>
    public string? Value { get; private set; }

    /// <summary>
    ///     Sets the string value (max length 255).
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    public IStringValueBuilder WithValue(string value)
    {
        ArgumentNullException.ThrowIfNull(value, nameof(value));
        EnsureNotBuilt(nameof(WithValue));
        Value = value;

        return this;
    }

    /// <summary>
    ///     Validates that the required properties of the builder have been set before building.
    /// </summary>
    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<StringValueBuilder, StringValue>(() => (nameof(Value), Value is not null));

    /// <summary>
    ///     Creates the StringValue instance.
    /// </summary>
    /// <returns>A new StringValue instance.</returns>
    protected override StringValue BuildCore() =>
        new()
        {
            Value = Value!
        };

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    protected override void ResetCore() => Value = null;
}