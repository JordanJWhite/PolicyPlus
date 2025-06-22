using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

namespace PolicyPlus.Infrastructure.Parsers.Utilities;

/// <summary>
///     Provides utility methods for XML parsing.
/// </summary>
public static class XmlParserUtilities
{
    /// <summary>
    ///     Regex for matching string references.
    /// </summary>
    public static readonly Regex StringReferenceRegex = new(@"\$\(string\.(\w+)\)", RegexOptions.Compiled);

    /// <summary>
    ///     Regex for matching presentation references.
    /// </summary>
    public static readonly Regex PresentationReferenceRegex = new(@"\$\(presentation\.(\w+)\)", RegexOptions.Compiled);

    /// <summary>
    ///     Gets a required attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <returns>The attribute value.</returns>
    /// <exception cref="PolicyMissingAttributeException">Thrown when the attribute is missing.</exception>
    public static string GetRequiredAttribute(XmlReader reader, string attributeName)
    {
        var value = reader.GetAttribute(attributeName);

        if (string.IsNullOrEmpty(value))
        {
            if (reader is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            {
                throw new PolicyMissingAttributeException(
                    attributeName,
                    reader.LocalName,
                    lineInfo.LineNumber,
                    lineInfo.LinePosition
                );
            }

            throw new PolicyMissingAttributeException(attributeName, reader.LocalName);
        }

        return value;
    }

    /// <summary>
    ///     Gets an optional attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <returns>The attribute value or null if not present.</returns>
    public static string? GetOptionalAttribute(XmlReader reader, string attributeName)
    {
        var value = reader.GetAttribute(attributeName);

        return string.IsNullOrEmpty(value) ? null : value;
    }

    /// <summary>
    ///     Gets a boolean attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <param name="defaultValue">The default value if the attribute is not present.</param>
    /// <returns>The boolean value.</returns>
    /// <exception cref="PolicyInvalidValueException">Thrown when the value cannot be parsed as boolean.</exception>
    public static bool GetBooleanAttribute(XmlReader reader, string attributeName, bool defaultValue = false)
    {
        var value = reader.GetAttribute(attributeName);

        if (string.IsNullOrEmpty(value))
            return defaultValue;

        if (bool.TryParse(value, out var result))
            return result;

        throw new PolicyInvalidValueException(value, "boolean (true/false)");
    }

    /// <summary>
    ///     Gets a uint attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <param name="defaultValue">The default value if the attribute is not present.</param>
    /// <returns>The uint value.</returns>
    /// <exception cref="PolicyInvalidValueException">Thrown when the value cannot be parsed as uint.</exception>
    public static uint GetUIntAttribute(XmlReader reader, string attributeName, uint defaultValue = 0)
    {
        var value = reader.GetAttribute(attributeName);

        if (string.IsNullOrEmpty(value))
            return defaultValue;

        if (uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
            return result;

        throw new PolicyInvalidValueException(value, "unsigned integer");
    }

    /// <summary>
    ///     Gets a ulong attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <param name="defaultValue">The default value if the attribute is not present.</param>
    /// <returns>The ulong value.</returns>
    /// <exception cref="PolicyInvalidValueException">Thrown when the value cannot be parsed as ulong.</exception>
    public static ulong GetULongAttribute(XmlReader reader, string attributeName, ulong defaultValue = 0)
    {
        var value = reader.GetAttribute(attributeName);

        if (string.IsNullOrEmpty(value))
            return defaultValue;

        if (ulong.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
            return result;

        throw new PolicyInvalidValueException(value, "unsigned long integer");
    }

    /// <summary>
    ///     Gets an optional uint attribute value from the XML reader.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="attributeName">The attribute name.</param>
    /// <returns>The uint value or null if not present.</returns>
    /// <exception cref="PolicyInvalidValueException">Thrown when the value cannot be parsed as uint.</exception>
    public static uint? GetOptionalUIntAttribute(XmlReader reader, string attributeName)
    {
        var value = reader.GetAttribute(attributeName);

        if (string.IsNullOrEmpty(value))
            return null;

        if (uint.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var result))
            return result;

        throw new PolicyInvalidValueException(value, "unsigned integer");
    }

    /// <summary>
    ///     Reads the element content as string asynchronously.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>The element content.</returns>
    public static async Task<string> ReadElementContentAsStringAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        if (reader.IsEmptyElement)
            return string.Empty;

        var content = await reader.ReadElementContentAsStringAsync();
        cancellationToken.ThrowIfCancellationRequested();

        return content;
    }

    /// <summary>
    ///     Reads to the next element asynchronously.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>True if an element was found; otherwise, false.</returns>
    public static async Task<bool> ReadToNextElementAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        while (await reader.ReadAsync())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader.NodeType == XmlNodeType.Element)
                return true;

            if (reader.NodeType == XmlNodeType.EndElement)
                return false;
        }

        return false;
    }

    /// <summary>
    ///     Skips the current element asynchronously.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="cancellationToken">The cancellation token.</param>
    /// <returns>A task representing the asynchronous operation.</returns>
    public static async Task SkipAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        if (reader.NodeType != XmlNodeType.Element || reader.IsEmptyElement)
            return;

        var depth = reader.Depth;

        while (await reader.ReadAsync())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth)
                break;
        }
    }

    /// <summary>
    ///     Ensures the reader is positioned on an element with the specified local name.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="expectedLocalName">The expected element local name.</param>
    /// <exception cref="PolicyUnexpectedElementException">Thrown when the element doesn't match.</exception>
    public static void EnsureElement(XmlReader reader, string expectedLocalName)
    {
        if (reader.NodeType != XmlNodeType.Element || reader.LocalName != expectedLocalName)
        {
            if (reader is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            {
                throw new PolicyUnexpectedElementException(
                    reader.LocalName,
                    expectedLocalName,
                    lineInfo.LineNumber,
                    lineInfo.LinePosition
                );
            }

            throw new PolicyUnexpectedElementException(reader.LocalName, expectedLocalName);
        }
    }

    /// <summary>
    ///     Checks if the reader is positioned on an element with the specified local name.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="localName">The element local name to check.</param>
    /// <returns>True if the reader is on the specified element; otherwise, false.</returns>
    public static bool IsElement(XmlReader reader, string localName) =>
        reader.NodeType == XmlNodeType.Element && reader.LocalName == localName;

    /// <summary>
    ///     Gets the current line information if available.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <returns>A tuple containing line number and position, or null if not available.</returns>
    public static (int LineNumber, int LinePosition)? GetLineInfo(XmlReader reader)
    {
        if (reader is IXmlLineInfo lineInfo && lineInfo.HasLineInfo())
            return (lineInfo.LineNumber, lineInfo.LinePosition);

        return null;
    }

    /// <summary>
    ///     Creates a parsing exception with line information if available.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="message">The error message.</param>
    /// <returns>The parsing exception.</returns>
    public static PolicyParsingException CreateParsingException(XmlReader reader, string message)
    {
        var lineInfo = GetLineInfo(reader);

        if (lineInfo.HasValue)
            return new PolicyParsingException(message, lineInfo.Value.LineNumber, lineInfo.Value.LinePosition);

        return new PolicyParsingException(message);
    }

    /// <summary>
    ///     Creates a parsing exception with line information if available.
    /// </summary>
    /// <param name="reader">The XML reader.</param>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The inner exception.</param>
    /// <returns>The parsing exception.</returns>
    public static PolicyParsingException CreateParsingException(XmlReader reader, string message, Exception innerException)
    {
        var lineInfo = GetLineInfo(reader);

        if (lineInfo.HasValue)
            return new PolicyParsingException(message, lineInfo.Value.LineNumber, lineInfo.Value.LinePosition, innerException);

        return new PolicyParsingException(message, innerException);
    }

    /// <summary>
    ///     Validates that a string is safe for use as an XML attribute or element value.
    /// </summary>
    /// <param name="value">The value to validate.</param>
    /// <returns>True if the value is safe; otherwise, false.</returns>
    public static bool IsXmlSafe(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return true;

        // Check for control characters (except tab, CR, LF)
        foreach (var c in value)
        {
            if (char.IsControl(c) && c != '\t' && c != '\r' && c != '\n')
                return false;
        }

        return true;
    }

    /// <summary>
    ///     Sanitizes a string for safe XML usage.
    /// </summary>
    /// <param name="value">The value to sanitize.</param>
    /// <returns>The sanitized value.</returns>
    public static string SanitizeXmlString(string? value)
    {
        if (string.IsNullOrEmpty(value))
            return string.Empty;

        // Remove control characters except tab, CR, LF
        var sanitized = new StringBuilder(value.Length);

        foreach (var c in value)
        {
            if (!char.IsControl(c) || c == '\t' || c == '\r' || c == '\n')
                sanitized.Append(c);
        }

        return sanitized.ToString();
    }
}