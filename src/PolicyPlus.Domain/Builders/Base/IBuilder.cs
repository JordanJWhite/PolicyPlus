namespace PolicyPlus.Domain.Builders.Base;

/// <summary>
///     Base interface for all builders.
/// </summary>
/// <typeparam name="T">The type of object to build.</typeparam>
public interface IBuilder<out T>
{
    /// <summary>
    ///     Builds the object.
    /// </summary>
    /// <returns>The built object.</returns>
    T Build();

    /// <summary>
    ///     Resets the builder to its initial state.
    /// </summary>
    void Reset();
}