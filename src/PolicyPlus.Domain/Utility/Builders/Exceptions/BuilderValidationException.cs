using System.Text.Json.Serialization;

namespace PolicyPlus.Domain.Utility.Builders.Exceptions;

/// <summary>
///     Exception thrown when a property value fails validation in a builder.
/// </summary>
public class BuilderValidationException : BuilderException
{
    /// <summary>
    ///     Represents the template for error messages generated during validation failures.
    /// </summary>
    /// <remarks>
    ///     The template includes placeholders for the field name, the context (e.g., class or method),
    ///     and the specific validation error message. Use <c>string.Format</c> to populate the placeholders.
    /// </remarks>
    private const string ErrorMessageTemplate = "Validation failed for property '{0}' in {1}: {2}";

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderValidationException" /> class for serialization.
    /// </summary>
    [JsonConstructor]
    public BuilderValidationException(
        string  message,
        string  propertyName,
        object? invalidValue,
        string? validationRule,
        string? builderTypeName,
        string? targetTypeName
    )
        : base(message, builderTypeName, targetTypeName)
    {
        PropertyName   = propertyName;
        InvalidValue   = invalidValue;
        ValidationRule = validationRule;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderValidationException" /> class.
    /// </summary>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="validationRule">The validation rule description.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderValidationException(string propertyName, object? invalidValue, string validationRule, Type builderType, Type targetType)
        : base(string.Format(ErrorMessageTemplate, propertyName, builderType.Name, validationRule), builderType, targetType)
    {
        PropertyName   = propertyName;
        InvalidValue   = invalidValue;
        ValidationRule = validationRule;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderValidationException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="invalidValue">The invalid value.</param>
    public BuilderValidationException(string message, string propertyName, object? invalidValue)
        : base(message)
    {
        PropertyName   = propertyName;
        InvalidValue   = invalidValue;
        ValidationRule = null;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderValidationException" /> class.
    /// </summary>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="validationRule">The validation rule description.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    /// <param name="innerException">The inner exception.</param>
    public BuilderValidationException(
        string    propertyName,
        object?   invalidValue,
        string    validationRule,
        Type      builderType,
        Type      targetType,
        Exception innerException
    )
        : base(
            string.Format(ErrorMessageTemplate, propertyName, builderType.Name, validationRule),
            builderType,
            targetType,
            innerException
        )
    {
        PropertyName   = propertyName;
        InvalidValue   = invalidValue;
        ValidationRule = validationRule;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderValidationException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="innerException">The inner exception.</param>
    public BuilderValidationException(string message, string propertyName, object? invalidValue, Exception innerException)
        : base(message, innerException)
    {
        PropertyName   = propertyName;
        InvalidValue   = invalidValue;
        ValidationRule = null;
    }

    /// <summary>
    ///     Gets the name of the property that failed validation.
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    ///     Gets the invalid value.
    /// </summary>
    public object? InvalidValue { get; }

    /// <summary>
    ///     Gets the validation rule that was violated.
    /// </summary>
    public string? ValidationRule { get; }
}