using System.Text.Json.Serialization;

namespace PolicyPlus.Domain.Builders.Exceptions;

/// <summary>
///     Exception thrown when a builder state is invalid.
/// </summary>
public class BuilderInvalidStateException : BuilderException
{
    private const string ErrorMessageTemplate = "Invalid state in {0}. Current: {1}, Expected: {2}";

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderInvalidStateException" /> class for serialization.
    /// </summary>
    [JsonConstructor]
    public BuilderInvalidStateException(
        string  message,
        string  currentState,
        string  expectedState,
        string? builderTypeName,
        string? targetTypeName
    )
        : base(message, builderTypeName, targetTypeName)
    {
        CurrentState  = currentState;
        ExpectedState = expectedState;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderInvalidStateException" /> class.
    /// </summary>
    /// <param name="currentState">The current state description.</param>
    /// <param name="expectedState">The expected state description.</param>
    /// <param name="builderType">The type of builder.</param>
    /// <param name="targetType">The type of object being built.</param>
    public BuilderInvalidStateException(string currentState, string expectedState, Type builderType, Type targetType)
        : base(string.Format(ErrorMessageTemplate, builderType.Name, currentState, expectedState), builderType, targetType)
    {
        CurrentState  = currentState;
        ExpectedState = expectedState;
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="BuilderInvalidStateException" /> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="currentState">The current state description.</param>
    /// <param name="expectedState">The expected state description.</param>
    public BuilderInvalidStateException(string message, string currentState, string expectedState)
        : base(message)
    {
        CurrentState  = currentState;
        ExpectedState = expectedState;
    }

    /// <summary>
    ///     Gets the current state description.
    /// </summary>
    public string CurrentState { get; }

    /// <summary>
    ///     Gets the expected state description.
    /// </summary>
    public string ExpectedState { get; }
}