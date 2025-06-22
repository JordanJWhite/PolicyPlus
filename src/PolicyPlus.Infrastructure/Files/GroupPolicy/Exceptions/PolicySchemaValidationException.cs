namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Exception thrown when XML schema validation fails.
/// </summary>
public class PolicySchemaValidationException : PolicyParsingException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicySchemaValidationException" /> class.
    /// </summary>
    /// <param name="validationError">The schema validation error.</param>
    public PolicySchemaValidationException(string validationError)
        : base($"Schema validation failed: {validationError}") =>
        ValidationError = validationError;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicySchemaValidationException" /> class.
    /// </summary>
    /// <param name="validationError">The schema validation error.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    public PolicySchemaValidationException(string validationError, int lineNumber, int linePosition)
        : base($"Schema validation failed: {validationError}", lineNumber, linePosition) =>
        ValidationError = validationError;

    /// <summary>
    ///     Gets the schema validation error details.
    /// </summary>
    public string ValidationError { get; }
}