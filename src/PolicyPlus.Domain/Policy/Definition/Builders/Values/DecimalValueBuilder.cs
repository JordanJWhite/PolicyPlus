using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for DecimalValue.
/// </summary>
public interface IDecimalValueBuilder : IBuilder<DecimalValue>
{
    /// <summary>
    ///     Sets the decimal value.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    IDecimalValueBuilder WithValue(uint value);
}

/// <summary>
///     Builder for DecimalValue.
/// </summary>
public class DecimalValueBuilder : BuilderBase<DecimalValueBuilder, DecimalValue>, IDecimalValueBuilder
{
    /// <summary>
    ///     Gets the value. Null indicates the value has not been set.
    /// </summary>
    public uint? Value { get; private set; }

    /// <summary>
    ///     Sets the decimal value.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    public IDecimalValueBuilder WithValue(uint value)
    {
        EnsureNotBuilt(nameof(WithValue));
        Value = value;

        return this;
    }

    /// <summary>
    ///     Validates that the required properties of the builder have been set before building.
    /// </summary>
    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<DecimalValueBuilder, DecimalValue>(() => (nameof(Value), Value.HasValue));

    /// <summary>
    ///     Creates the DecimalValue instance.
    /// </summary>
    /// <returns>A new DecimalValue instance.</returns>
    protected override DecimalValue BuildCore() =>
        new()
        {
            Value = Value!.Value
        };

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    protected override void ResetCore() => Value = null;
}