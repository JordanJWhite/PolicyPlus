using System.Text.Json.Serialization;

namespace PolicyPlus.Domain.Builders.Exceptions;

/// <summary>
///     Base exception for all builder-related errors.
/// </summary>
public class BuilderException : Exception
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class.
    /// </summary>
    public BuilderException() { }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public BuilderException(string message)
        : base(message) { }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The inner exception.</param>
    public BuilderException(string message, Exception innerException)
        : base(message, innerException) { }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class for serialization.
    /// </summary>
    [JsonConstructor]
    public BuilderException(string message, string? builderTypeName, string? targetTypeName)
        : base(message)
    {
        BuilderTypeName = builderTypeName;
        TargetTypeName  = targetTypeName;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderException(string message, Type builderType, Type targetType)
        : base(message)
    {
        BuilderType     = builderType;
        TargetType      = targetType;
        BuilderTypeName = builderType.Name;
        TargetTypeName  = targetType.Name;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    /// <param name="innerException">The inner exception.</param>
    public BuilderException(string message, Type builderType, Type targetType, Exception innerException)
        : base(message, innerException)
    {
        BuilderType     = builderType;
        TargetType      = targetType;
        BuilderTypeName = builderType.Name;
        TargetTypeName  = targetType.Name;
    }

    /// <summary>
    ///     Gets the type of builder that threw the exception. Not serialized.
    /// </summary>
    [JsonIgnore]
    public Type? BuilderType { get; }

    /// <summary>
    ///     Gets the type of object being built. Not serialized.
    /// </summary>
    [JsonIgnore]
    public Type? TargetType { get; }

    /// <summary>
    ///     Gets the name of the builder type that threw the exception.
    /// </summary>
    public string? BuilderTypeName { get; }

    /// <summary>
    ///     Gets the name of the type of object being built.
    /// </summary>
    public string? TargetTypeName { get; }
}