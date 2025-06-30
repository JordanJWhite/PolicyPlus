using PolicyPlus.Domain.Policy.Definition.Model.Values;
using PolicyPlus.Domain.Utility.Builders.Base;
using PolicyPlus.Domain.Utility.Builders.Helpers;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Values;

/// <summary>
///     Builder interface for ValueList.
/// </summary>
public interface IValueListBuilder : IBuilder<ValueList>
{
    IValueListBuilder AddItem(ValueItem                  item);
    IValueListBuilder AddItems(IEnumerable<ValueItem>    items);
    IValueListBuilder RemoveItem(ValueItem               item);
    IValueListBuilder RemoveItems(IEnumerable<ValueItem> items);
    IValueListBuilder ClearItems();
    IValueListBuilder WithDefaultKey(string? defaultKey);
}

/// <summary>
///     Builder for ValueList.
/// </summary>
public class ValueListBuilder : BuilderBase<ValueListBuilder, ValueList>, IValueListBuilder
{
    private List<ValueItem>           _items = [];
    private IReadOnlyList<ValueItem>? _itemsView;
    private bool                      _itemsSet;

    public IReadOnlyList<ValueItem> Items      => _itemsView ??= _items.AsReadOnly();
    public string?                  DefaultKey { get; private set; }

    public IValueListBuilder AddItem(ValueItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));

        EnsureNotBuilt(nameof(AddItem));

        _items.Add(item);
        _itemsSet = true;

        return this;
    }

    public IValueListBuilder AddItems(IEnumerable<ValueItem> items)
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

    public IValueListBuilder RemoveItem(ValueItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(RemoveItem));

        _items.Remove(item);
        _itemsSet = _items.Count > 0;

        return this;
    }

    public IValueListBuilder RemoveItems(IEnumerable<ValueItem> items)
    {
        ArgumentNullException.ThrowIfNull(items, nameof(items));
        EnsureNotBuilt(nameof(RemoveItems));

        foreach (var item in items)
        {
            if (item == null)
                throw new ArgumentNullException(nameof(items), "Items collection cannot contain null values.");

            _items.Remove(item);
        }

        _itemsSet = _items.Count > 0;

        return this;
    }

    public IValueListBuilder ClearItems()
    {
        EnsureNotBuilt(nameof(ClearItems));

        _items.Clear();
        _itemsSet = false;

        return this;
    }

    public IValueListBuilder WithDefaultKey(string? defaultKey)
    {
        EnsureNotBuilt(nameof(WithDefaultKey));
        DefaultKey = defaultKey;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<ValueListBuilder, ValueList>(() => (nameof(Items), _itemsSet));

    protected override ValueList BuildCore() =>
        new()
        {
            Items      = _items.AsReadOnly(),
            DefaultKey = DefaultKey
        };

    protected override void ResetCore()
    {
        DefaultKey = null;
        _items     = [];
        _itemsView = null;
        _itemsSet  = false;
    }
}