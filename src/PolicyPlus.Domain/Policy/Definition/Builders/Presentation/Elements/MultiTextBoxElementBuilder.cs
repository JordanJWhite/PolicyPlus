using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for MultiTextBoxElement.
/// </summary>
public interface IMultiTextBoxElementBuilder : IBuilder<MultiTextBoxElement>
{
    IMultiTextBoxElementBuilder WithLabel(string       label);
    IMultiTextBoxElementBuilder WithShowAsDialog(bool  showAsDialog);
    IMultiTextBoxElementBuilder WithDefaultHeight(uint defaultHeight);
    IMultiTextBoxElementBuilder WithRefId(string       refId);
}

/// <summary>
///     Builder for MultiTextBoxElement.
/// </summary>
public class MultiTextBoxElementBuilder : BuilderBase<MultiTextBoxElementBuilder, MultiTextBoxElement>, IMultiTextBoxElementBuilder
{
    public string? Label         { get; private set; }
    public bool    ShowAsDialog  { get; private set; }
    public uint    DefaultHeight { get; private set; } = 3;
    public string? RefId         { get; private set; }

    public IMultiTextBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public IMultiTextBoxElementBuilder WithShowAsDialog(bool showAsDialog)
    {
        EnsureNotBuilt(nameof(WithShowAsDialog));
        ShowAsDialog = showAsDialog;

        return this;
    }

    public IMultiTextBoxElementBuilder WithDefaultHeight(uint defaultHeight)
    {
        EnsureNotBuilt(nameof(WithDefaultHeight));
        DefaultHeight = defaultHeight;

        return this;
    }

    public IMultiTextBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<MultiTextBoxElementBuilder, MultiTextBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override MultiTextBoxElement BuildCore() =>
        new()
        {
            Label         = Label!,
            ShowAsDialog  = ShowAsDialog,
            DefaultHeight = DefaultHeight,
            RefId         = RefId!
        };

    protected override void ResetCore()
    {
        Label         = null;
        ShowAsDialog  = false;
        DefaultHeight = 3;
        RefId         = null;
    }
}