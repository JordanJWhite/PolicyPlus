using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for LongDecimalTextBoxElement.
/// </summary>
public interface ILongDecimalTextBoxElementBuilder : IBuilder<LongDecimalTextBoxElement>
{
    ILongDecimalTextBoxElementBuilder WithLabel(string       label);
    ILongDecimalTextBoxElementBuilder WithDefaultValue(ulong defaultValue);
    ILongDecimalTextBoxElementBuilder WithSpin(bool          spin);
    ILongDecimalTextBoxElementBuilder WithSpinStep(uint      spinStep);
    ILongDecimalTextBoxElementBuilder WithRefId(string       refId);
}

/// <summary>
///     Builder for LongDecimalTextBoxElement.
/// </summary>
public class LongDecimalTextBoxElementBuilder
    : BuilderBase<LongDecimalTextBoxElementBuilder, LongDecimalTextBoxElement>, ILongDecimalTextBoxElementBuilder
{
    public string? Label        { get; private set; }
    public ulong   DefaultValue { get; private set; } = 1;
    public bool    Spin         { get; private set; } = true;
    public uint    SpinStep     { get; private set; } = 1;
    public string? RefId        { get; private set; }

    public ILongDecimalTextBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public ILongDecimalTextBoxElementBuilder WithDefaultValue(ulong defaultValue)
    {
        EnsureNotBuilt(nameof(WithDefaultValue));
        DefaultValue = defaultValue;

        return this;
    }

    public ILongDecimalTextBoxElementBuilder WithSpin(bool spin)
    {
        EnsureNotBuilt(nameof(WithSpin));
        Spin = spin;

        return this;
    }

    public ILongDecimalTextBoxElementBuilder WithSpinStep(uint spinStep)
    {
        EnsureNotBuilt(nameof(WithSpinStep));
        SpinStep = spinStep;

        return this;
    }

    public ILongDecimalTextBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<LongDecimalTextBoxElementBuilder, LongDecimalTextBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override LongDecimalTextBoxElement BuildCore() =>
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