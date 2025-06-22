namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Exception thrown when an unexpected element is encountered.
/// </summary>
public class PolicyUnexpectedElementException : PolicyParsingException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyUnexpectedElementException" /> class.
    /// </summary>
    /// <param name="unexpectedElement">The unexpected element name.</param>
    /// <param name="parentElement">The parent element name.</param>
    public PolicyUnexpectedElementException(string unexpectedElement, string parentElement)
        : base($"Unexpected element '{unexpectedElement}' found in '{parentElement}'")
    {
        UnexpectedElement = unexpectedElement;
        ParentElement     = parentElement;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyUnexpectedElementException" /> class.
    /// </summary>
    /// <param name="unexpectedElement">The unexpected element name.</param>
    /// <param name="parentElement">The parent element name.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    public PolicyUnexpectedElementException(string unexpectedElement, string parentElement, int lineNumber, int linePosition)
        : base($"Unexpected element '{unexpectedElement}' found in '{parentElement}'", lineNumber, linePosition)
    {
        UnexpectedElement = unexpectedElement;
        ParentElement     = parentElement;
    }

    /// <summary>
    ///     Gets the unexpected element name.
    /// </summary>
    public string UnexpectedElement { get; }

    /// <summary>
    ///     Gets the parent element name.
    /// </summary>
    public string ParentElement { get; }
}