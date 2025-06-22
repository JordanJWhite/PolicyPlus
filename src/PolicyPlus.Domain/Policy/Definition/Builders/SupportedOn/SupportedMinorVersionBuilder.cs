using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedMinorVersion.
/// </summary>
public interface ISupportedMinorVersionBuilder : IBuilder<SupportedMinorVersion>
{
    ISupportedMinorVersionBuilder WithName(string name);
    ISupportedMinorVersionBuilder WithDisplayName(string displayName);
    ISupportedMinorVersionBuilder WithVersionIndex(uint versionIndex);
}

/// <summary>
///     Builder for SupportedMinorVersion.
/// </summary>
public class SupportedMinorVersionBuilder : BuilderBase<SupportedMinorVersionBuilder, SupportedMinorVersion>, ISupportedMinorVersionBuilder
{
    private bool _nameSet;
    private bool _displayNameSet;
    private bool _versionIndexSet;

    public string? Name { get; private set; }
    public string? DisplayName { get; private set; }
    public uint? VersionIndex { get; private set; }

    public ISupportedMinorVersionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));
        
        Name = name;
        _nameSet = true;

        return this;
    }

    public ISupportedMinorVersionBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));
        
        DisplayName = displayName;
        _displayNameSet = true;

        return this;
    }

    public ISupportedMinorVersionBuilder WithVersionIndex(uint versionIndex)
    {
        EnsureNotBuilt(nameof(WithVersionIndex));
        
        VersionIndex = versionIndex;
        _versionIndexSet = true;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedMinorVersionBuilder, SupportedMinorVersion>(
            () => (nameof(Name), _nameSet),
            () => (nameof(DisplayName), _displayNameSet),
            () => (nameof(VersionIndex), _versionIndexSet)
        );

    protected override SupportedMinorVersion BuildCore() =>
        new()
        {
            Name = Name!,
            DisplayName = DisplayName!,
            VersionIndex = VersionIndex!.Value
        };

    protected override void ResetCore()
    {
        Name = null;
        DisplayName = null;
        VersionIndex = null;
        _nameSet = false;
        _displayNameSet = false;
        _versionIndexSet = false;
    }
}