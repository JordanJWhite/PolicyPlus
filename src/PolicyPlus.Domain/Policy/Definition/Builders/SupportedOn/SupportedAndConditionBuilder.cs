using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedAndCondition.
/// </summary>
public interface ISupportedAndConditionBuilder : IBuilder<SupportedAndCondition>
{
    ISupportedAndConditionBuilder AddItem(ISupportedOnItem                  item);
    ISupportedAndConditionBuilder AddItems(IEnumerable<ISupportedOnItem>    items);
    ISupportedAndConditionBuilder RemoveItem(ISupportedOnItem               item);
    ISupportedAndConditionBuilder RemoveItems(IEnumerable<ISupportedOnItem> items);
    ISupportedAndConditionBuilder ClearItems();
}

/// <summary>
///     Builder for SupportedAndCondition.
/// </summary>
public class SupportedAndConditionBuilder : BuilderBase<SupportedAndConditionBuilder, SupportedAndCondition>, ISupportedAndConditionBuilder
{
    private List<ISupportedOnItem>           _items = [];
    private IReadOnlyList<ISupportedOnItem>? _itemsView;
    private bool                             _itemsSet;

    public IReadOnlyList<ISupportedOnItem> Items => _itemsView ??= _items.AsReadOnly();

    public ISupportedAndConditionBuilder AddItem(ISupportedOnItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(AddItem));

        _items.Add(item);
        _itemsSet = true;

        return this;
    }

    public ISupportedAndConditionBuilder AddItems(IEnumerable<ISupportedOnItem> items)
    {
        ArgumentNullException.ThrowIfNull(items, nameof(items));
        EnsureNotBuilt(nameof(AddItems));

        foreach (var item in items)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(items), "Items collection cannot contain null values.");

            _items.Add(item);
        }

        _itemsSet = true;

        return this;
    }

    public ISupportedAndConditionBuilder RemoveItem(ISupportedOnItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(RemoveItem));

        _items.Remove(item);

        return this;
    }

    public ISupportedAndConditionBuilder RemoveItems(IEnumerable<ISupportedOnItem> items)
    {
        ArgumentNullException.ThrowIfNull(items, nameof(items));
        EnsureNotBuilt(nameof(RemoveItems));

        foreach (var item in items)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(items), "Items collection cannot contain null values.");

            _items.Remove(item);
        }

        return this;
    }

    public ISupportedAndConditionBuilder ClearItems()
    {
        EnsureNotBuilt(nameof(ClearItems));
        _items.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<SupportedAndConditionBuilder, SupportedAndCondition>(() => (nameof(Items), _itemsSet));

    protected override SupportedAndCondition BuildCore() =>
        new()
        {
            Items = Items
        };

    protected override void ResetCore()
    {
        _items     = [];
        _itemsView = null;
        _itemsSet  = false;
    }
}