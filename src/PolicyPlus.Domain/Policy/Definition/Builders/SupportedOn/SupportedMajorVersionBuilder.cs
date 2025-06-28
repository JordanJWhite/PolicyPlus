using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedMajorVersion.
/// </summary>
public interface ISupportedMajorVersionBuilder : IBuilder<SupportedMajorVersion>
{
    ISupportedMajorVersionBuilder WithName(string                                        name);
    ISupportedMajorVersionBuilder WithDisplayName(string                                 displayName);
    ISupportedMajorVersionBuilder WithVersionIndex(uint                                  versionIndex);
    ISupportedMajorVersionBuilder AddMinorVersion(SupportedMinorVersion                  minorVersion);
    ISupportedMajorVersionBuilder AddMinorVersions(IEnumerable<SupportedMinorVersion>    minorVersions);
    ISupportedMajorVersionBuilder RemoveMinorVersion(SupportedMinorVersion               minorVersion);
    ISupportedMajorVersionBuilder RemoveMinorVersions(IEnumerable<SupportedMinorVersion> minorVersions);
    ISupportedMajorVersionBuilder ClearMinorVersions();
}

/// <summary>
///     Builder for SupportedMajorVersion.
/// </summary>
public class SupportedMajorVersionBuilder : BuilderBase<SupportedMajorVersionBuilder, SupportedMajorVersion>, ISupportedMajorVersionBuilder
{
    private List<SupportedMinorVersion>           _minorVersions = [];
    private IReadOnlyList<SupportedMinorVersion>? _minorVersionsView;

    public string?                              Name          { get; private set; }
    public string?                              DisplayName   { get; private set; }
    public uint?                                VersionIndex  { get; private set; }
    public IReadOnlyList<SupportedMinorVersion> MinorVersions => _minorVersionsView ??= _minorVersions.AsReadOnly();

    public ISupportedMajorVersionBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));

        Name = name;

        return this;
    }

    public ISupportedMajorVersionBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName = displayName;

        return this;
    }

    public ISupportedMajorVersionBuilder WithVersionIndex(uint versionIndex)
    {
        EnsureNotBuilt(nameof(WithVersionIndex));

        VersionIndex = versionIndex;

        return this;
    }

    public ISupportedMajorVersionBuilder AddMinorVersion(SupportedMinorVersion minorVersion)
    {
        ArgumentNullException.ThrowIfNull(minorVersion, nameof(minorVersion));
        EnsureNotBuilt(nameof(AddMinorVersion));

        _minorVersions.Add(minorVersion);

        return this;
    }

    public ISupportedMajorVersionBuilder AddMinorVersions(IEnumerable<SupportedMinorVersion> minorVersions)
    {
        ArgumentNullException.ThrowIfNull(minorVersions, nameof(minorVersions));
        EnsureNotBuilt(nameof(AddMinorVersions));

        foreach (var minorVersion in minorVersions)
        {
            if (minorVersion == null)
                throw new ArgumentNullException(nameof(minorVersions), "Minor versions collection cannot contain null values.");

            _minorVersions.Add(minorVersion);
        }

        return this;
    }

    public ISupportedMajorVersionBuilder RemoveMinorVersion(SupportedMinorVersion minorVersion)
    {
        ArgumentNullException.ThrowIfNull(minorVersion, nameof(minorVersion));
        EnsureNotBuilt(nameof(RemoveMinorVersion));

        _minorVersions.Remove(minorVersion);

        return this;
    }

    public ISupportedMajorVersionBuilder RemoveMinorVersions(IEnumerable<SupportedMinorVersion> minorVersions)
    {
        ArgumentNullException.ThrowIfNull(minorVersions, nameof(minorVersions));
        EnsureNotBuilt(nameof(RemoveMinorVersions));

        foreach (var minorVersion in minorVersions)
        {
            if (minorVersion == null)
                throw new ArgumentNullException(nameof(minorVersions), "Minor versions collection cannot contain null values.");

            _minorVersions.Remove(minorVersion);
        }

        return this;
    }

    public ISupportedMajorVersionBuilder ClearMinorVersions()
    {
        EnsureNotBuilt(nameof(ClearMinorVersions));
        _minorVersions.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedMajorVersionBuilder, SupportedMajorVersion>(
            () => (nameof(Name), Name is not null),
            () => (nameof(DisplayName), DisplayName is not null),
            () => (nameof(VersionIndex), VersionIndex is not null)
        );

    protected override SupportedMajorVersion BuildCore() =>
        new()
        {
            Name          = Name!,
            DisplayName   = DisplayName!,
            VersionIndex  = VersionIndex!.Value,
            MinorVersions = MinorVersions
        };

    protected override void ResetCore()
    {
        Name               = null;
        DisplayName        = null;
        VersionIndex       = null;
        _minorVersions     = [];
        _minorVersionsView = null;
    }
}