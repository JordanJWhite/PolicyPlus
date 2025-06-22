using System.Numerics;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Model.Elements;

/// <summary>
///     Represents a boolean (checkbox) element in a policy.
/// </summary>
public record BooleanElement
    : PolicyElementBase,
      IEqualityOperators<BooleanElement, BooleanElement, bool>
{
    /// <summary>
    ///     Gets the value when true.
    /// </summary>
    public IValue? TrueValue { get; init; }

    /// <summary>
    ///     Gets the value when false.
    /// </summary>
    public IValue? FalseValue { get; init; }

    /// <summary>
    ///     Gets the value list when true.
    /// </summary>
    public ValueList? TrueList { get; init; }

    /// <summary>
    ///     Gets the value list when false.
    /// </summary>
    public ValueList? FalseList { get; init; }
}