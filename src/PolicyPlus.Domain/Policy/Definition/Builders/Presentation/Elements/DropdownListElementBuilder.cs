using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for DropdownListElement.
/// </summary>
public interface IDropdownListElementBuilder : IBuilder<DropdownListElement>
{
    IDropdownListElementBuilder WithLabel(string      label);
    IDropdownListElementBuilder WithNoSort(bool       noSort);
    IDropdownListElementBuilder WithDefaultItem(uint? defaultItem);
    IDropdownListElementBuilder WithRefId(string      refId);
}

/// <summary>
///     Builder for DropdownListElement.
/// </summary>
public class DropdownListElementBuilder : BuilderBase<DropdownListElementBuilder, DropdownListElement>, IDropdownListElementBuilder
{
    public string? Label       { get; private set; }
    public bool    NoSort      { get; private set; }
    public uint?   DefaultItem { get; private set; }
    public string? RefId       { get; private set; }

    public IDropdownListElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public IDropdownListElementBuilder WithNoSort(bool noSort)
    {
        EnsureNotBuilt(nameof(WithNoSort));
        NoSort = noSort;

        return this;
    }

    public IDropdownListElementBuilder WithDefaultItem(uint? defaultItem)
    {
        EnsureNotBuilt(nameof(WithDefaultItem));
        DefaultItem = defaultItem;

        return this;
    }

    public IDropdownListElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<DropdownListElementBuilder, DropdownListElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override DropdownListElement BuildCore() =>
        new()
        {
            Label       = Label!,
            NoSort      = NoSort,
            DefaultItem = DefaultItem,
            RefId       = RefId!
        };

    protected override void ResetCore()
    {
        Label       = null;
        NoSort      = false;
        DefaultItem = null;
        RefId       = null;
    }
}