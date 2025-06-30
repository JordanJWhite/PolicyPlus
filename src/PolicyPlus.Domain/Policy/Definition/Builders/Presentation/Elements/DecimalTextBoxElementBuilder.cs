using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for DecimalTextBoxElement.
/// </summary>
public interface IDecimalTextBoxElementBuilder : IBuilder<DecimalTextBoxElement>
{
    IDecimalTextBoxElementBuilder WithLabel(string      label);
    IDecimalTextBoxElementBuilder WithDefaultValue(uint defaultValue);
    IDecimalTextBoxElementBuilder WithSpin(bool         spin);
    IDecimalTextBoxElementBuilder WithSpinStep(uint     spinStep);
    IDecimalTextBoxElementBuilder WithRefId(string      refId);
}

/// <summary>
///     Builder for DecimalTextBoxElement.
/// </summary>
public class DecimalTextBoxElementBuilder : BuilderBase<DecimalTextBoxElementBuilder, DecimalTextBoxElement>, IDecimalTextBoxElementBuilder
{
    public string? Label        { get; private set; }
    public uint    DefaultValue { get; private set; } = 1;
    public bool    Spin         { get; private set; } = true;
    public uint    SpinStep     { get; private set; } = 1;
    public string? RefId        { get; private set; }

    public IDecimalTextBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public IDecimalTextBoxElementBuilder WithDefaultValue(uint defaultValue)
    {
        EnsureNotBuilt(nameof(WithDefaultValue));
        DefaultValue = defaultValue;

        return this;
    }

    public IDecimalTextBoxElementBuilder WithSpin(bool spin)
    {
        EnsureNotBuilt(nameof(WithSpin));
        Spin = spin;

        return this;
    }

    public IDecimalTextBoxElementBuilder WithSpinStep(uint spinStep)
    {
        EnsureNotBuilt(nameof(WithSpinStep));
        SpinStep = spinStep;

        return this;
    }

    public IDecimalTextBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<DecimalTextBoxElementBuilder, DecimalTextBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override DecimalTextBoxElement BuildCore() =>
        new()
        {
            Label        = Label!,
            DefaultValue = DefaultValue,
            Spin         = Spin,
            SpinStep     = SpinStep,
            RefId        = RefId!
        };

    protected override void ResetCore()
    {
        Label        = null;
        DefaultValue = 1;
        Spin         = true;
        SpinStep     = 1;
        RefId        = null;
    }
}