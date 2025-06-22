using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for DeleteValue.
/// </summary>
public interface IDeleteValueBuilder : IBuilder<DeleteValue>
{
    // DeleteValue has no properties to set
}

/// <summary>
///     Builder for DeleteValue.
/// </summary>
public class DeleteValueBuilder : BuilderBase<DeleteValueBuilder, DeleteValue>, IDeleteValueBuilder
{
    /// <summary>
    ///     Validates that the required properties of the builder have been set before building.
    /// </summary>
    protected override void ValidateRequiredProperties()
    {
        // DeleteValue has no required properties
    }

    /// <summary>
    ///     Creates the DeleteValue instance.
    /// </summary>
    /// <returns>A new DeleteValue instance.</returns>
    protected override DeleteValue BuildCore() => new();

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    protected override void ResetCore()
    {
        // DeleteValue has no properties to reset
    }
}