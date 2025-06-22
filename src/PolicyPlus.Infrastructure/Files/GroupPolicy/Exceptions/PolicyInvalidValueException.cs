namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Exceptions;

/// <summary>
///     Exception thrown when an invalid value is encountered.
/// </summary>
public class PolicyInvalidValueException : PolicyParsingException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyInvalidValueException" /> class.
    /// </summary>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="expectedFormat">The expected type or format.</param>
    public PolicyInvalidValueException(string invalidValue, string expectedFormat)
        : base($"Invalid value '{invalidValue}'. Expected: {expectedFormat}")
    {
        InvalidValue   = invalidValue;
        ExpectedFormat = expectedFormat;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="PolicyInvalidValueException" /> class.
    /// </summary>
    /// <param name="invalidValue">The invalid value.</param>
    /// <param name="expectedFormat">The expected type or format.</param>
    /// <param name="innerException">The inner exception.</param>
    public PolicyInvalidValueException(string invalidValue, string expectedFormat, Exception innerException)
        : base($"Invalid value '{invalidValue}'. Expected: {expectedFormat}", innerException)
    {
        InvalidValue   = invalidValue;
        ExpectedFormat = expectedFormat;
    }

    /// <summary>
    ///     Gets the invalid value.
    /// </summary>
    public string InvalidValue { get; }

    /// <summary>
    ///     Gets the expected type or format.
    /// </summary>
    public string ExpectedFormat { get; }
}