using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnReference.
/// </summary>
public interface ISupportedOnReferenceBuilder : IBuilder<SupportedOnReference>
{
    ISupportedOnReferenceBuilder WithRef(string reference);
}

/// <summary>
///     Builder for SupportedOnReference.
/// </summary>
public class SupportedOnReferenceBuilder : BuilderBase<SupportedOnReferenceBuilder, SupportedOnReference>, ISupportedOnReferenceBuilder
{
    public string? Ref { get; private set; }

    public ISupportedOnReferenceBuilder WithRef(string reference)
    {
        ArgumentNullException.ThrowIfNull(reference, nameof(reference));
        EnsureNotBuilt(nameof(WithRef));

        Ref = reference;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<SupportedOnReferenceBuilder, SupportedOnReference>(() => (nameof(Ref), Ref is not null));

    protected override SupportedOnReference BuildCore() =>
        new()
        {
            Ref = Ref!
        };

    protected override void ResetCore() => Ref = null;
}