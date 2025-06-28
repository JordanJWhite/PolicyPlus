using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnReferenceItem.
/// </summary>
public interface ISupportedOnReferenceItemBuilder : IBuilder<SupportedOnReferenceItem>
{
    ISupportedOnReferenceItemBuilder WithReference(SupportedOnReference reference);
}

/// <summary>
///     Builder for SupportedOnReferenceItem.
/// </summary>
public class SupportedOnReferenceItemBuilder
    : BuilderBase<SupportedOnReferenceItemBuilder, SupportedOnReferenceItem>, ISupportedOnReferenceItemBuilder
{
    public SupportedOnReference? Reference { get; private set; }

    public ISupportedOnReferenceItemBuilder WithReference(SupportedOnReference reference)
    {
        EnsureNotBuilt(nameof(WithReference));
        Reference = reference;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<SupportedOnReferenceItemBuilder, SupportedOnReferenceItem>(() => (nameof(Reference),
                                                                                                               Reference is not null)
            );

    protected override SupportedOnReferenceItem BuildCore() =>
        new()
        {
            Reference = Reference!.Value
        };

    protected override void ResetCore() => Reference = null;
}