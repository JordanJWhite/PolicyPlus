namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Base exception for all policy parsing related errors.
/// </summary>
public class PolicyParsingException : Exception
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyParsingException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public PolicyParsingException(string message)
        : base(message) { }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyParsingException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The inner exception.</param>
    public PolicyParsingException(string message, Exception innerException)
        : base(message, innerException) { }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyParsingException" /> class with location information.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    public PolicyParsingException(string message, int lineNumber, int linePosition)
        : base($"{message} (Line: {lineNumber}, Position: {linePosition})")
    {
        LineNumber   = lineNumber;
        LinePosition = linePosition;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyParsingException" /> class with location information.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    /// <param name="innerException">The inner exception.</param>
    public PolicyParsingException(string message, int lineNumber, int linePosition, Exception innerException)
        : base($"{message} (Line: {lineNumber}, Position: {linePosition})", innerException)
    {
        LineNumber   = lineNumber;
        LinePosition = linePosition;
    }

    /// <summary>
    ///     Gets the line number where the error occurred.
    /// </summary>
    public int? LineNumber { get; }

    /// <summary>
    ///     Gets the column position where the error occurred.
    /// </summary>
    public int? LinePosition { get; }
}