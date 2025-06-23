using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedOnDefinition.
/// </summary>
public interface ISupportedOnDefinitionBuilder : IBuilder<SupportedOnDefinition>
{
    ISupportedOnDefinitionBuilder WithName(string                      name);
    ISupportedOnDefinitionBuilder WithDisplayName(string               displayName);
    ISupportedOnDefinitionBuilder WithCondition(ISupportedOnCondition? condition);
}

/// <summary>
///     Builder for SupportedOnDefinition.
/// </summary>
public class SupportedOnDefinitionBuilder : BuilderBase<SupportedOnDefinitionBuilder, SupportedOnDefinition>, ISupportedOnDefinitionBuilder
{
    private bool _nameSet;
    private bool _displayNameSet;

    public string?                Name        { get; private set; }
    public string?                DisplayName { get; private set; }
    public ISupportedOnCondition? Condition   { get; private set; }

    public ISupportedOnDefinitionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));

        Name     = name;
        _nameSet = true;

        return this;
    }

    public ISupportedOnDefinitionBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName     = displayName;
        _displayNameSet = true;

        return this;
    }

    public ISupportedOnDefinitionBuilder WithCondition(ISupportedOnCondition? condition)
    {
        EnsureNotBuilt(nameof(WithCondition));
        Condition = condition;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedOnDefinitionBuilder, SupportedOnDefinition>(
            () => (nameof(Name), _nameSet),
            () => (nameof(DisplayName), _displayNameSet)
        );

    protected override SupportedOnDefinition BuildCore() =>
        new()
        {
            Name        = Name!,
            DisplayName = DisplayName!,
            Condition   = Condition
        };

    protected override void ResetCore()
    {
        Name            = null;
        DisplayName     = null;
        Condition       = null;
        _nameSet        = false;
        _displayNameSet = false;
    }
}