using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;
using PolicyPlus.Domain.Utility.Builders.Base;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnTable.
/// </summary>
public interface ISupportedOnTableBuilder : IBuilder<SupportedOnTable>
{
    ISupportedOnTableBuilder AddProduct(SupportedProduct                  product);
    ISupportedOnTableBuilder AddProducts(IEnumerable<SupportedProduct>    products);
    ISupportedOnTableBuilder RemoveProduct(SupportedProduct               product);
    ISupportedOnTableBuilder RemoveProducts(IEnumerable<SupportedProduct> products);
    ISupportedOnTableBuilder ClearProducts();
    ISupportedOnTableBuilder AddDefinition(SupportedOnDefinition                  definition);
    ISupportedOnTableBuilder AddDefinitions(IEnumerable<SupportedOnDefinition>    definitions);
    ISupportedOnTableBuilder RemoveDefinition(SupportedOnDefinition               definition);
    ISupportedOnTableBuilder RemoveDefinitions(IEnumerable<SupportedOnDefinition> definitions);
    ISupportedOnTableBuilder ClearDefinitions();
}

/// <summary>
///     Builder for SupportedOnTable.
/// </summary>
public class SupportedOnTableBuilder : BuilderBase<SupportedOnTableBuilder, SupportedOnTable>, ISupportedOnTableBuilder
{
    // Fields to hold the products and definitions
    private List<SupportedProduct>      _products    = [];
    private List<SupportedOnDefinition> _definitions = [];

    // Views for the products and definitions.
    private IReadOnlyList<SupportedProduct>?      _productsView;
    private IReadOnlyList<SupportedOnDefinition>? _definitionsView;

    public IReadOnlyList<SupportedProduct>      Products    => _productsView ??= _products.AsReadOnly();
    public IReadOnlyList<SupportedOnDefinition> Definitions => _definitionsView ??= _definitions.AsReadOnly();

    public ISupportedOnTableBuilder AddProduct(SupportedProduct product)
    {
        ArgumentNullException.ThrowIfNull(product, nameof(product));
        EnsureNotBuilt(nameof(AddProduct));

        _products.Add(product);

        return this;
    }

    public ISupportedOnTableBuilder AddProducts(IEnumerable<SupportedProduct> products)
    {
        ArgumentNullException.ThrowIfNull(products, nameof(products));
        EnsureNotBuilt(nameof(AddProducts));

        foreach (var product in products)
        {
            if (product == null)
                throw new ArgumentNullException(nameof(products), "Products collection cannot contain null values.");

            _products.Add(product);
        }

        return this;
    }

    public ISupportedOnTableBuilder RemoveProduct(SupportedProduct product)
    {
        ArgumentNullException.ThrowIfNull(product, nameof(product));
        EnsureNotBuilt(nameof(RemoveProduct));

        _products.Remove(product);

        return this;
    }

    public ISupportedOnTableBuilder RemoveProducts(IEnumerable<SupportedProduct> products)
    {
        ArgumentNullException.ThrowIfNull(products, nameof(products));
        EnsureNotBuilt(nameof(RemoveProducts));

        foreach (var product in products)
        {
            if (product == null)
                throw new ArgumentNullException(nameof(products), "Products collection cannot contain null values.");

            _products.Remove(product);
        }

        return this;
    }

    public ISupportedOnTableBuilder ClearProducts()
    {
        EnsureNotBuilt(nameof(ClearProducts));
        _products.Clear();

        return this;
    }

    public ISupportedOnTableBuilder AddDefinition(SupportedOnDefinition definition)
    {
        ArgumentNullException.ThrowIfNull(definition, nameof(definition));
        EnsureNotBuilt(nameof(AddDefinition));

        _definitions.Add(definition);

        return this;
    }

    public ISupportedOnTableBuilder AddDefinitions(IEnumerable<SupportedOnDefinition> definitions)
    {
        ArgumentNullException.ThrowIfNull(definitions, nameof(definitions));
        EnsureNotBuilt(nameof(AddDefinitions));

        foreach (var definition in definitions)
        {
            if (definition == null)
                throw new ArgumentNullException(nameof(definitions), "Definitions collection cannot contain null values.");

            _definitions.Add(definition);
        }

        return this;
    }

    public ISupportedOnTableBuilder RemoveDefinition(SupportedOnDefinition definition)
    {
        ArgumentNullException.ThrowIfNull(definition, nameof(definition));
        EnsureNotBuilt(nameof(RemoveDefinition));

        _definitions.Remove(definition);

        return this;
    }

    public ISupportedOnTableBuilder RemoveDefinitions(IEnumerable<SupportedOnDefinition> definitions)
    {
        ArgumentNullException.ThrowIfNull(definitions, nameof(definitions));
        EnsureNotBuilt(nameof(RemoveDefinitions));

        foreach (var definition in definitions)
        {
            if (definition == null)
                throw new ArgumentNullException(nameof(definitions), "Definitions collection cannot contain null values.");

            _definitions.Remove(definition);
        }

        return this;
    }

    public ISupportedOnTableBuilder ClearDefinitions()
    {
        EnsureNotBuilt(nameof(ClearDefinitions));
        _definitions.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties()
    {
        // SupportedOnTable has no required properties
    }

    protected override SupportedOnTable BuildCore() =>
        new()
        {
            Products    = Products,
            Definitions = Definitions
        };

    protected override void ResetCore()
    {
        _products        = [];
        _definitions     = [];
        _productsView    = null;
        _definitionsView = null;
    }
}