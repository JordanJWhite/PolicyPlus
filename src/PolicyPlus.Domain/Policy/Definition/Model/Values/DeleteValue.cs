using System.Numerics;

namespace PolicyPlus.Domain.Policy.Definition.Model.Values;

/// <summary>
///     Represents a delete value operation.
/// </summary>
public record struct DeleteValue
    : IValue,
      IEqualityOperators<DeleteValue, DeleteValue, bool>;