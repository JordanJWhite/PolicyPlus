using System.Text.Json.Serialization;

namespace PolicyPlus.Domain.Builders.Exceptions;

/// <summary>
///     Exception thrown when a required property is not set in a builder.
/// </summary>
public class BuilderRequiredPropertyException : BuilderException
{
    private const string ErrorMessageSinglePropertyTemplate     = "Required property '{0}' is not set in {1} for {2}";
    private const string ErrorMessageMultiplePropertiesTemplate = "Required properties are not set in {0} for {1}: {2}";

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderRequiredPropertyException" /> class for serialization.
    /// </summary>
    [JsonConstructor]
    public BuilderRequiredPropertyException(
        string                message,
        IReadOnlyList<string> missingProperties,
        string                builderTypeName,
        string                targetTypeName
    )
        : base(message, builderTypeName, targetTypeName) =>
        MissingProperties = missingProperties;

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderRequiredPropertyException" /> class for a single missing
    ///     required property.
    /// </summary>
    /// <param name="missingProperty">The name of the missing required property.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderRequiredPropertyException(string missingProperty, Type builderType, Type targetType)
        : base(string.Format(ErrorMessageSinglePropertyTemplate, missingProperty, builderType.Name, targetType.Name)) =>
        MissingProperties = [missingProperty];

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderRequiredPropertyException" /> class for multiple missing
    ///     required properties.
    /// </summary>
    /// <param name="missingProperties">The list of missing required properties.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderRequiredPropertyException(IEnumerable<string> missingProperties, Type builderType, Type targetType)
        : base(
            string.Format(ErrorMessageMultiplePropertiesTemplate, builderType.Name, targetType.Name, string.Join(", ", missingProperties)),
            builderType,
            targetType
        ) =>
        MissingProperties = missingProperties.ToList()
                                             .AsReadOnly();

    /// <summary>
    ///     Gets the names of the missing properties.
    /// </summary>
    public IReadOnlyList<string> MissingProperties { get; }
}