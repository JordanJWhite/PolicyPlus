using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Localization;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Localization;

/// <summary>
///     Builder interface for LocalizedString.
/// </summary>
public interface ILocalizedStringBuilder : IBuilder<LocalizedString>
{
    ILocalizedStringBuilder WithId(string    id);
    ILocalizedStringBuilder WithValue(string value);
}

/// <summary>
///     Builder for LocalizedString.
/// </summary>
public class LocalizedStringBuilder : BuilderBase<LocalizedStringBuilder, LocalizedString>, ILocalizedStringBuilder
{
    public string? Id    { get; private set; }
    public string? Value { get; private set; }

    public ILocalizedStringBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));

        Id = id;

        return this;
    }

    public ILocalizedStringBuilder WithValue(string value)
    {
        ArgumentNullException.ThrowIfNull(value, nameof(value));
        EnsureNotBuilt(nameof(WithValue));

        Value = value;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<LocalizedStringBuilder, LocalizedString>(
            () => (nameof(Id), Id is not null),
            () => (nameof(Value), Value is not null)
        );

    protected override LocalizedString BuildCore() =>
        new()
        {
            Id    = Id!,
            Value = Value!
        };

    protected override void ResetCore()
    {
        Id    = null;
        Value = null;
    }
}