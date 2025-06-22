using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.References;

/// <summary>
///     Represents a reference to a file.
/// </summary>
public record struct FileReference
    : IEqualityOperators<FileReference, FileReference, bool>
{
    /// <summary>
    ///     Gets the file name.
    /// </summary>
    public required string FileName { get; init; }
}