using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for EnumerationElement.
/// </summary>
public interface IEnumerationElementBuilder : IBuilder<EnumerationElement>
{
    IEnumerationElementBuilder WithId(string                            id);
    IEnumerationElementBuilder WithClientExtension(string?              clientExtension);
    IEnumerationElementBuilder WithKey(string?                          key);
    IEnumerationElementBuilder WithValueName(string?                    valueName);
    IEnumerationElementBuilder WithRequired(bool                        required);
    IEnumerationElementBuilder AddItem(EnumerationItem                  item);
    IEnumerationElementBuilder AddItems(IEnumerable<EnumerationItem>    items);
    IEnumerationElementBuilder RemoveItem(EnumerationItem               item);
    IEnumerationElementBuilder RemoveItems(IEnumerable<EnumerationItem> items);
    IEnumerationElementBuilder ClearItems();
}

/// <summary>
///     Builder for EnumerationElement.
/// </summary>
public class EnumerationElementBuilder : BuilderBase<EnumerationElementBuilder, EnumerationElement>, IEnumerationElementBuilder
{
    private List<EnumerationItem>           _items = [];
    private IReadOnlyList<EnumerationItem>? _itemsView;
    private bool                            _itemsSet;

    public string?                        Id              { get; private set; }
    public string?                        ClientExtension { get; private set; }
    public string?                        Key             { get; private set; }
    public string?                        ValueName       { get; private set; }
    public bool                           Required        { get; private set; }
    public IReadOnlyList<EnumerationItem> Items           => _itemsView ??= _items.AsReadOnly();

    public IEnumerationElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));
        Id = id;

        return this;
    }

    public IEnumerationElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public IEnumerationElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IEnumerationElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public IEnumerationElementBuilder WithRequired(bool required)
    {
        EnsureNotBuilt(nameof(WithRequired));
        Required = required;

        return this;
    }

    public IEnumerationElementBuilder AddItem(EnumerationItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(AddItem));

        _items.Add(item);
        _itemsSet = true;

        return this;
    }

    public IEnumerationElementBuilder AddItems(IEnumerable<EnumerationItem> items)
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

    public IEnumerationElementBuilder RemoveItem(EnumerationItem item)
    {
        ArgumentNullException.ThrowIfNull(item, nameof(item));
        EnsureNotBuilt(nameof(RemoveItem));

        _items.Remove(item);
        _itemsSet = _items.Count > 0;

        return this;
    }

    public IEnumerationElementBuilder RemoveItems(IEnumerable<EnumerationItem> items)
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

    public IEnumerationElementBuilder ClearItems()
    {
        EnsureNotBuilt(nameof(ClearItems));

        _items.Clear();
        _itemsSet = false;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<EnumerationElementBuilder, EnumerationElement>(
                () => (nameof(Id), Id is not null),
                () => (nameof(Items), _itemsSet)
            );

    protected override EnumerationElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValueName       = ValueName,
            Required        = Required,
            Items           = _items.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Id              = null;
        ClientExtension = null;
        Key             = null;
        ValueName       = null;
        Required        = false;
        _items          = [];
        _itemsView      = null;
        _itemsSet       = false;
    }
}