using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Core;

/// <summary>
///     Represents application-specific data embedded in an ADMX or ADML file.
///     This allows tools to store custom metadata within the policy files.
/// </summary>
public record struct Annotation
    : IEqualityOperators<Annotation, Annotation, bool>
{
    /// <summary>
    ///     Gets the identifier for the application that this annotation targets.
    /// </summary>
    public required string Application { get; init; }

    /// <summary>
    ///     Gets the raw, arbitrary XML or text content of the annotation.
    ///     The consuming application is responsible for interpreting this data.
    /// </summary>
    public required string Content { get; init; }
}