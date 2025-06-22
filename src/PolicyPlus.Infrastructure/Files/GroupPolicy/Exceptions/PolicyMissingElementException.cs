namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Exception thrown when a required element is missing.
/// </summary>
public class PolicyMissingElementException : PolicyParsingException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyMissingElementException" /> class.
    /// </summary>
    /// <param name="elementName">The name of the missing element.</param>
    public PolicyMissingElementException(string elementName)
        : base($"Required element '{elementName}' is missing") =>
        ElementName = elementName;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyMissingElementException" /> class.
    /// </summary>
    /// <param name="elementName">The name of the missing element.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    public PolicyMissingElementException(string elementName, int lineNumber, int linePosition)
        : base($"Required element '{elementName}' is missing", lineNumber, linePosition) =>
        ElementName = elementName;

    /// <summary>
    ///     Gets the name of the missing element.
    /// </summary>
    public string ElementName { get; }
}