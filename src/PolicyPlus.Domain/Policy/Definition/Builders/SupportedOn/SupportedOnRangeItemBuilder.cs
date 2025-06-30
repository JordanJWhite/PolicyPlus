using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;
using PolicyPlus.Domain.Utility.Builders.Base;
using PolicyPlus.Domain.Utility.Builders.Helpers;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnRangeItem.
/// </summary>
public interface ISupportedOnRangeItemBuilder : IBuilder<SupportedOnRangeItem>
{
    ISupportedOnRangeItemBuilder WithRange(SupportedOnRange range);
}

/// <summary>
///     Builder for SupportedOnRangeItem.
/// </summary>
public class SupportedOnRangeItemBuilder : BuilderBase<SupportedOnRangeItemBuilder, SupportedOnRangeItem>, ISupportedOnRangeItemBuilder
{
    public SupportedOnRange? Range { get; private set; }

    public ISupportedOnRangeItemBuilder WithRange(SupportedOnRange range)
    {
        EnsureNotBuilt(nameof(WithRange));
        Range = range;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedOnRangeItemBuilder, SupportedOnRangeItem>(() => (nameof(Range),
                                                                                                                          Range is not null)
        );

    protected override SupportedOnRangeItem BuildCore() =>
        new()
        {
            Range = Range!.Value
        };

    protected override void ResetCore() => Range = null;
}