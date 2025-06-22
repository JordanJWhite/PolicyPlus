namespace PolicyPlus.Domain.Policy.Definition.Model.Core;

/// <summary>
///     Represents the class of policy definition.
/// </summary>
public enum PolicyClass
{
    /// <summary>
    ///     User-specific policy.
    /// </summary>
    User,

    /// <summary>
    ///     Machine-specific policy.
    /// </summary>
    Machine,

    /// <summary>
    ///     Policy applies to both user and machine.
    /// </summary>
    Both
}