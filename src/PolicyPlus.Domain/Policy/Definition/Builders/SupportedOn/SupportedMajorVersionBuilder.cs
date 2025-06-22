using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedMajorVersion.
/// </summary>
public interface ISupportedMajorVersionBuilder : IBuilder<SupportedMajorVersion>
{
    ISupportedMajorVersionBuilder WithName(string name);
    ISupportedMajorVersionBuilder WithDisplayName(string displayName);
    ISupportedMajorVersionBuilder WithVersionIndex(uint versionIndex);
    ISupportedMajorVersionBuilder AddMinorVersion(SupportedMinorVersion minorVersion);
}

/// <summary>
///     Builder for SupportedMajorVersion.
/// </summary>
public class SupportedMajorVersionBuilder : BuilderBase<SupportedMajorVersionBuilder, SupportedMajorVersion>, ISupportedMajorVersionBuilder
{
    private readonly List<SupportedMinorVersion> _minorVersions = new();
    private bool _nameSet;
    private bool _displayNameSet;
    private bool _versionIndexSet;
    
    public string? Name { get; private set; }
    public string? DisplayName { get; private set; }
    public uint? VersionIndex { get; private set; }
    public IReadOnlyList<SupportedMinorVersion> MinorVersions => _minorVersions.AsReadOnly();

    public ISupportedMajorVersionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));
        
        Name = name;
        _nameSet = true;

        return this;
    }

    public ISupportedMajorVersionBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));
        
        DisplayName = displayName;
        _displayNameSet = true;

        return this;
    }

    public ISupportedMajorVersionBuilder WithVersionIndex(uint versionIndex)
    {
        EnsureNotBuilt(nameof(WithVersionIndex));
        
        VersionIndex = versionIndex;
        _versionIndexSet = true;

        return this;
    }

    public ISupportedMajorVersionBuilder AddMinorVersion(SupportedMinorVersion minorVersion)
    {
        ArgumentNullException.ThrowIfNull(minorVersion, nameof(minorVersion));
        EnsureNotBuilt(nameof(AddMinorVersion));
        
        _minorVersions.Add(minorVersion);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedMajorVersionBuilder, SupportedMajorVersion>(
            () => (nameof(Name), _nameSet),
            () => (nameof(DisplayName), _displayNameSet),
            () => (nameof(VersionIndex), _versionIndexSet)
        );

    protected override SupportedMajorVersion BuildCore() =>
        new()
        {
            Name = Name!,
            DisplayName = DisplayName!,
            VersionIndex = VersionIndex!.Value,
            MinorVersions = _minorVersions.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Name = null;
        DisplayName = null;
        VersionIndex = null;
        _minorVersions.Clear();
        _nameSet = false;
        _displayNameSet = false;
        _versionIndexSet = false;
    }
}