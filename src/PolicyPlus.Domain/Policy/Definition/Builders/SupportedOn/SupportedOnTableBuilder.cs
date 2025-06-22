using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnTable.
/// </summary>
public interface ISupportedOnTableBuilder : IBuilder<SupportedOnTable>
{
    ISupportedOnTableBuilder AddProduct(SupportedProduct         product);
    ISupportedOnTableBuilder AddDefinition(SupportedOnDefinition definition);
}

/// <summary>
///     Builder for SupportedOnTable.
/// </summary>
public class SupportedOnTableBuilder : BuilderBase<SupportedOnTableBuilder, SupportedOnTable>, ISupportedOnTableBuilder
{
    private readonly List<SupportedOnDefinition> _definitions = new();
    private readonly List<SupportedProduct>      _products    = new();

    public IReadOnlyList<SupportedProduct>      Products    => _products.AsReadOnly();
    public IReadOnlyList<SupportedOnDefinition> Definitions => _definitions.AsReadOnly();

    public ISupportedOnTableBuilder AddProduct(SupportedProduct product)
    {
        ArgumentNullException.ThrowIfNull(product, nameof(product));
        EnsureNotBuilt(nameof(AddProduct));

        _products.Add(product);

        return this;
    }

    public ISupportedOnTableBuilder AddDefinition(SupportedOnDefinition definition)
    {
        ArgumentNullException.ThrowIfNull(definition, nameof(definition));
        EnsureNotBuilt(nameof(AddDefinition));

        _definitions.Add(definition);

        return this;
    }

    protected override void ValidateRequiredProperties()
    {
        // SupportedOnTable has no required properties
    }

    protected override SupportedOnTable BuildCore() =>
        new()
        {
            Products    = _products.AsReadOnly(),
            Definitions = _definitions.AsReadOnly()
        };

    protected override void ResetCore()
    {
        _products.Clear();
        _definitions.Clear();
    }
}