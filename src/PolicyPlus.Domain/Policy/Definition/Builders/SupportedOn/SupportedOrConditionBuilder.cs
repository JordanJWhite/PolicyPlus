using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;
using PolicyPlus.Domain.Utility.Builders.Base;
using PolicyPlus.Domain.Utility.Builders.Helpers;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOrCondition.
/// </summary>
public interface ISupportedOrConditionBuilder : IBuilder<SupportedOrCondition>
{
    ISupportedOrConditionBuilder AddItem(ISupportedOnItem                  item);
    ISupportedOrConditionBuilder AddItems(IEnumerable<ISupportedOnItem>    items);
    ISupportedOrConditionBuilder RemoveItem(ISupportedOnItem               item);
    ISupportedOrConditionBuilder RemoveItems(IEnumerable<ISupportedOnItem> items);
    ISupportedOrConditionBuilder ClearItems();
}

/// <summary>
///     Builder for SupportedOrCondition.
/// </summary>
public class SupportedOrConditionBuilder : BuilderBase<SupportedOrConditionBuilder, SupportedOrCondition>, ISupportedOrConditionBuilder
{
    private List<ISupportedOnItem>           _items = [];
    private IReadOnlyList<ISupportedOnItem>? _itemsView;
    private bool                             _itemsSet;

    public IReadOnlyList<ISupportedOnItem> Items => _itemsView ??= _items.AsReadOnly();

    public ISupportedOrConditionBuilder AddItem(ISupportedOnItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(AddItem));

        _items.Add(item);
        _itemsSet = true;

        return this;
    }

    public ISupportedOrConditionBuilder AddItems(IEnumerable<ISupportedOnItem> items)
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

    public ISupportedOrConditionBuilder RemoveItem(ISupportedOnItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(RemoveItem));

        _items.Remove(item);

        return this;
    }

    public ISupportedOrConditionBuilder RemoveItems(IEnumerable<ISupportedOnItem> items)
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

    public ISupportedOrConditionBuilder ClearItems()
    {
        EnsureNotBuilt(nameof(ClearItems));

        _items.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedOrConditionBuilder, SupportedOrCondition>(() => (nameof(Items),
                                                                                                                          _itemsSet)
        );

    protected override SupportedOrCondition BuildCore() =>
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