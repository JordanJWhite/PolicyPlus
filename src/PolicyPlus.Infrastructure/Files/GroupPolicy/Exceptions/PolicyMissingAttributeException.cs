namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Exception thrown when a required attribute is missing.
/// </summary>
public class PolicyMissingAttributeException : PolicyParsingException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyMissingAttributeException" /> class.
    /// </summary>
    /// <param name="attributeName">The name of the missing attribute.</param>
    /// <param name="elementName">The element name containing the missing attribute.</param>
    public PolicyMissingAttributeException(string attributeName, string elementName)
        : base($"Required attribute '{attributeName}' is missing from element '{elementName}'")
    {
        AttributeName = attributeName;
        ElementName   = elementName;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyMissingAttributeException" /> class.
    /// </summary>
    /// <param name="attributeName">The name of the missing attribute.</param>
    /// <param name="elementName">The element name containing the missing attribute.</param>
    /// <param name="lineNumber">The line number where the error occurred.</param>
    /// <param name="linePosition">The column position where the error occurred.</param>
    public PolicyMissingAttributeException(string attributeName, string elementName, int lineNumber, int linePosition)
        : base($"Required attribute '{attributeName}' is missing from element '{elementName}'", lineNumber, linePosition)
    {
        AttributeName = attributeName;
        ElementName   = elementName;
    }

    /// <summary>
    ///     Gets the name of the missing attribute.
    /// </summary>
    public string AttributeName { get; }

    /// <summary>
    ///     Gets the element name containing the missing attribute.
    /// </summary>
    public string ElementName { get; }
}