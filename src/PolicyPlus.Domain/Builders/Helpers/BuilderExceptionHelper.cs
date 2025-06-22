using PolicyPlus.Domain.Builders.Exceptions;

namespace PolicyPlus.Domain.Builders.Helpers;

/// <summary>
///     Helper class for creating builder exceptions with consistent messaging.
/// </summary>
public static class BuilderExceptionHelper
{
    /// <summary>
    ///     Throws a required property exception for missing required properties.
    /// </summary>
    /// <typeparam name="TBuilder">The builder type.</typeparam>
    /// <typeparam name="TTarget">The target type being built.</typeparam>
    /// <param name="properties">The required properties to be validated.</param>
    /// <exception cref="BuilderRequiredPropertyException">Thrown when any required property has not been set.</exception>
    public static void ThrowIfRequiredPropertiesMissing<TBuilder, TTarget>(
        params IEnumerable<Func<(string PropertyName, bool propertyWasSet)>> properties
    )
    {
        var propertiesUnpacked = properties.Select(property => property())
                                           .ToList();

        var anyUnset = propertiesUnpacked.Any(property => !property.propertyWasSet);

        if (!anyUnset)
            return;

        var propertyNames = propertiesUnpacked.Where(property => !property.propertyWasSet)
                                              .Select(property => property.PropertyName);

        throw new BuilderRequiredPropertyException(
            propertyNames,
            typeof(TBuilder),
            typeof(TTarget)
        );
    }

    /// <summary>
    ///     Creates a validation exception.
    /// </summary>
    /// <typeparam name="TBuilder">The builder type.</typeparam>
    /// <typeparam name="TTarget">The target type being built.</typeparam>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="rule">The validation rule description.</param>
    /// <returns>A new BuilderValidationException.</returns>
    public static BuilderValidationException Validation<TBuilder, TTarget>(string propertyName, object? invalidValue, string rule) =>
        new(propertyName, invalidValue, rule, typeof(TBuilder), typeof(TTarget));

    /// <summary>
    ///     Creates an already built exception.
    /// </summary>
    /// <typeparam name="TBuilder">The builder type.</typeparam>
    /// <typeparam name="TTarget">The target type being built.</typeparam>
    /// <returns>A new BuilderAlreadyBuiltException.</returns>
    public static BuilderAlreadyBuiltException ThrowAlreadyBuilt<TBuilder, TTarget>() => new(typeof(TBuilder), typeof(TTarget));

    /// <summary>
    ///     Creates an already built exception with method name.
    /// </summary>
    /// <typeparam name="TBuilder">The builder type.</typeparam>
    /// <typeparam name="TTarget">The target type being built.</typeparam>
    /// <param name="methodName">The method that was attempted.</param>
    /// <returns>A new BuilderAlreadyBuiltException.</returns>
    public static BuilderAlreadyBuiltException ThrowAlreadyBuilt<TBuilder, TTarget>(string methodName) =>
        new(methodName, typeof(TBuilder), typeof(TTarget));
}