using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnRange.
/// </summary>
public interface ISupportedOnRangeBuilder : IBuilder<SupportedOnRange>
{
    ISupportedOnRangeBuilder WithRef(string            reference);
    ISupportedOnRangeBuilder WithMinVersionIndex(uint? minVersionIndex);
    ISupportedOnRangeBuilder WithMaxVersionIndex(uint? maxVersionIndex);
}

/// <summary>
///     Builder for SupportedOnRange.
/// </summary>
public class SupportedOnRangeBuilder : BuilderBase<SupportedOnRangeBuilder, SupportedOnRange>, ISupportedOnRangeBuilder
{
    public string? Ref             { get; private set; }
    public uint?   MinVersionIndex { get; private set; }
    public uint?   MaxVersionIndex { get; private set; }

    public ISupportedOnRangeBuilder WithRef(string reference)
    {
        ArgumentNullException.ThrowIfNull(reference, nameof(reference));
        EnsureNotBuilt(nameof(WithRef));

        Ref = reference;

        return this;
    }

    public ISupportedOnRangeBuilder WithMinVersionIndex(uint? minVersionIndex)
    {
        EnsureNotBuilt(nameof(WithMinVersionIndex));
        MinVersionIndex = minVersionIndex;

        return this;
    }

    public ISupportedOnRangeBuilder WithMaxVersionIndex(uint? maxVersionIndex)
    {
        EnsureNotBuilt(nameof(WithMaxVersionIndex));
        MaxVersionIndex = maxVersionIndex;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<SupportedOnRangeBuilder, SupportedOnRange>(() => (nameof(Ref), Ref is not null));

    protected override SupportedOnRange BuildCore() =>
        new()
        {
            Ref             = Ref!,
            MinVersionIndex = MinVersionIndex,
            MaxVersionIndex = MaxVersionIndex
        };

    protected override void ResetCore()
    {
        Ref             = null;
        MinVersionIndex = null;
        MaxVersionIndex = null;
    }
}