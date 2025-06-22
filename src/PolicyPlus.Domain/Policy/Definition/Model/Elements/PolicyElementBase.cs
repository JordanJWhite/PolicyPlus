using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Base class for all policy elements.
/// </summary>
public abstract record PolicyElementBase
    : IEqualityOperators<PolicyElementBase, PolicyElementBase, bool>
{
    /// <summary>
    ///     Gets the unique identifier for this element.
    /// </summary>
    public required string Id { get; init; }

    /// <summary>
    ///     Gets the optional client extension GUID.
    /// </summary>
    public string? ClientExtension { get; init; }

    /// <summary>
    ///     Gets the optional registry key.
    /// </summary>
    public string? Key { get; init; }

    /// <summary>
    ///     Gets the optional registry value name.
    /// </summary>
    public string? ValueName { get; init; }
}