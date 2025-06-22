using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for LongDecimalValue.
/// </summary>
public interface ILongDecimalValueBuilder : IBuilder<LongDecimalValue>
{
    /// <summary>
    ///     Sets the long decimal value.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    ILongDecimalValueBuilder WithValue(ulong value);
}

/// <summary>
///     Builder for LongDecimalValue.
/// </summary>
public class LongDecimalValueBuilder : BuilderBase<LongDecimalValueBuilder, LongDecimalValue>, ILongDecimalValueBuilder
{
    /// <summary>
    ///     Gets the value. Null indicates the value has not been set.
    /// </summary>
    public ulong? Value { get; private set; }

    /// <summary>
    ///     Sets the long decimal value.
    /// </summary>
    /// <param name="value">The value to set.</param>
    /// <returns>The builder instance for method chaining.</returns>
    public ILongDecimalValueBuilder WithValue(ulong value)
    {
        EnsureNotBuilt(nameof(WithValue));
        Value = value;

        return this;
    }

    /// <summary>
    ///     Validates that the required properties of the builder have been set before building.
    /// </summary>
    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<LongDecimalValueBuilder, LongDecimalValue>(() => (nameof(Value),
                                                                                                                  Value.HasValue)
        );

    /// <summary>
    ///     Creates the LongDecimalValue instance.
    /// </summary>
    /// <returns>A new LongDecimalValue instance.</returns>
    protected override LongDecimalValue BuildCore() =>
        new()
        {
            Value = Value!.Value
        };

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    protected override void ResetCore() => Value = null;
}