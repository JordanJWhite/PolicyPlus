using PolicyPlus.Domain.Builders.Helpers;

namespace PolicyPlus.Domain.Builders.Base;

/// <summary>
///     Base implementation for builders.
/// </summary>
/// <typeparam name="TChild">The type of the child class of <see cref="BuilderBase{TChild,TBuildObject}" /></typeparam>
/// <typeparam name="TBuildObject">The type of object to build.</typeparam>
/// <remarks>Thread Safety: This class is NOT thread-safe. Each thread should use its own builder instance.</remarks>
public abstract class BuilderBase<TChild, TBuildObject> : IBuilder<TBuildObject>
{
    /// <summary>
    ///     Gets a value indicating whether the builder has been built.
    /// </summary>
    protected bool IsBuilt { get; private set; }

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    public virtual void Reset()
    {
        IsBuilt = false;
        ResetCore();
    }

    /// <summary>
    ///     Builds the object.
    /// </summary>
    /// <returns>The built object.</returns>
    public TBuildObject Build()
    {
        if (IsBuilt)
            BuilderExceptionHelper.ThrowAlreadyBuilt<TChild, TBuildObject>(nameof(Build));

        ValidateRequiredProperties();
        var result = BuildCore();
        IsBuilt = true;

        return result;
    }

    /// <summary>
    ///     Validates that the required properties of the builder have been set before building.
    /// </summary>
    protected abstract void ValidateRequiredProperties();

    /// <summary>
    ///     Core build implementation.
    /// </summary>
    /// <returns>The built object.</returns>
    protected abstract TBuildObject BuildCore();

    /// <summary>
    ///     Core reset implementation.
    /// </summary>
    protected abstract void ResetCore();

    /// <summary>
    ///     Ensures the builder has not been built.
    /// </summary>
    /// <param name="callingMethodName">The name of the method calling this method.</param>
    protected void EnsureNotBuilt(string callingMethodName)
    {
        if (IsBuilt)
            BuilderExceptionHelper.ThrowAlreadyBuilt<TChild, TBuildObject>(callingMethodName);
    }
}