using System.Text.Json.Serialization;

namespace PolicyPlus.Domain.Utility.Builders.Exceptions;

/// <summary>
///     Exception thrown when attempting to modify a builder after it has been built.
/// </summary>
public class BuilderAlreadyBuiltException : BuilderException
{
    private const string ErrorMessageNoMethodNameTemplate = "Cannot modify {0} after it has been built. Call Reset() to reuse the builder.";

    private const string ErrorMessageWithMethodNameTemplate =
        "Cannot call {0} on {1} after it has been built. Call Reset() to reuse the builder.";

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderAlreadyBuiltException" /> class for serialization.
    /// </summary>
    [JsonConstructor]
    public BuilderAlreadyBuiltException(string message, string attemptedMethod, string builderTypeName, string targetTypeName)
        : base(message, builderTypeName, targetTypeName) =>
        AttemptedMethod = attemptedMethod;

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderAlreadyBuiltException" /> class.
    /// </summary>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderAlreadyBuiltException(Type builderType, Type targetType)
        : base(string.Format(ErrorMessageNoMethodNameTemplate, builderType.Name), builderType, targetType) =>
        AttemptedMethod = string.Empty;

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderAlreadyBuiltException" /> class.
    /// </summary>
    /// <param name="attemptedMethod">The method that was attempted.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderAlreadyBuiltException(string attemptedMethod, Type builderType, Type targetType)
        : base(
            string.Format(ErrorMessageWithMethodNameTemplate, attemptedMethod, builderType.Name),
            builderType,
            targetType
        ) =>
        AttemptedMethod = attemptedMethod;

    /// <summary>
    ///     Gets the method that was called on the built builder.
    /// </summary>
    public string AttemptedMethod { get; }
}