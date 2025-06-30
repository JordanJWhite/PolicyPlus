using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for ListElement.
/// </summary>
public interface IListElementBuilder : IBuilder<ListElement>
{
    IListElementBuilder WithId(string               id);
    IListElementBuilder WithClientExtension(string? clientExtension);
    IListElementBuilder WithKey(string?             key);
    IListElementBuilder WithValuePrefix(string?     valuePrefix);
    IListElementBuilder WithAdditive(bool           additive);
    IListElementBuilder WithExpandable(bool         expandable);
    IListElementBuilder WithExplicitValue(bool      explicitValue);
}

/// <summary>
///     Builder for ListElement.
/// </summary>
public class ListElementBuilder : BuilderBase<ListElementBuilder, ListElement>, IListElementBuilder
{
    public string? Id              { get; private set; }
    public string? ClientExtension { get; private set; }
    public string? Key             { get; private set; }
    public string? ValuePrefix     { get; private set; }
    public bool    Additive        { get; private set; }
    public bool    Expandable      { get; private set; }
    public bool    ExplicitValue   { get; private set; }

    public IListElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));
        Id = id;

        return this;
    }

    public IListElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public IListElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IListElementBuilder WithValuePrefix(string? valuePrefix)
    {
        EnsureNotBuilt(nameof(WithValuePrefix));
        ValuePrefix = valuePrefix;

        return this;
    }

    public IListElementBuilder WithAdditive(bool additive)
    {
        EnsureNotBuilt(nameof(WithAdditive));
        Additive = additive;

        return this;
    }

    public IListElementBuilder WithExpandable(bool expandable)
    {
        EnsureNotBuilt(nameof(WithExpandable));
        Expandable = expandable;

        return this;
    }

    public IListElementBuilder WithExplicitValue(bool explicitValue)
    {
        EnsureNotBuilt(nameof(WithExplicitValue));
        ExplicitValue = explicitValue;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<ListElementBuilder, ListElement>(() => (nameof(Id), Id is not null));

    protected override ListElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValuePrefix     = ValuePrefix,
            Additive        = Additive,
            Expandable      = Expandable,
            ExplicitValue   = ExplicitValue
        };

    protected override void ResetCore()
    {
        Id              = null;
        ClientExtension = null;
        Key             = null;
        ValuePrefix     = null;
        Additive        = false;
        Expandable      = false;
        ExplicitValue   = false;
    }
}