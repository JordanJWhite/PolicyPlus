using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Localization;

/// <summary>
///     Builder interface for LocalizationResourceReference.
/// </summary>
public interface ILocalizationResourceReferenceBuilder : IBuilder<LocalizationResourceReference>
{
    ILocalizationResourceReferenceBuilder WithMinRequiredRevision(string minRequiredRevision);
    ILocalizationResourceReferenceBuilder WithFallbackCulture(string     fallbackCulture);
}

/// <summary>
///     Builder for LocalizationResourceReference.
/// </summary>
public class LocalizationResourceReferenceBuilder
    : BuilderBase<LocalizationResourceReferenceBuilder, LocalizationResourceReference>, ILocalizationResourceReferenceBuilder
{
    public string? MinRequiredRevision { get; private set; }
    public string  FallbackCulture     { get; private set; } = "en-US";

    public ILocalizationResourceReferenceBuilder WithMinRequiredRevision(string minRequiredRevision)
    {
        ArgumentNullException.ThrowIfNull(minRequiredRevision, nameof(minRequiredRevision));
        EnsureNotBuilt(nameof(WithMinRequiredRevision));

        MinRequiredRevision = minRequiredRevision;

        return this;
    }

    public ILocalizationResourceReferenceBuilder WithFallbackCulture(string fallbackCulture)
    {
        ArgumentNullException.ThrowIfNull(fallbackCulture, nameof(fallbackCulture));
        EnsureNotBuilt(nameof(WithFallbackCulture));

        FallbackCulture = fallbackCulture;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<LocalizationResourceReferenceBuilder,
                LocalizationResourceReference>(() => (nameof(MinRequiredRevision), MinRequiredRevision is not null));

    protected override LocalizationResourceReference BuildCore() =>
        new()
        {
            MinRequiredRevision = MinRequiredRevision!,
            FallbackCulture     = FallbackCulture
        };

    protected override void ResetCore()
    {
        MinRequiredRevision = null;
        FallbackCulture     = "en-US";
    }
}