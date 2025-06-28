using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedProduct.
/// </summary>
public interface ISupportedProductBuilder : IBuilder<SupportedProduct>
{
    ISupportedProductBuilder WithName(string                                        name);
    ISupportedProductBuilder WithDisplayName(string                                 displayName);
    ISupportedProductBuilder AddMajorVersion(SupportedMajorVersion                  majorVersion);
    ISupportedProductBuilder AddMajorVersions(IEnumerable<SupportedMajorVersion>    majorVersions);
    ISupportedProductBuilder RemoveMajorVersion(SupportedMajorVersion               majorVersion);
    ISupportedProductBuilder RemoveMajorVersions(IEnumerable<SupportedMajorVersion> majorVersions);
    ISupportedProductBuilder ClearMajorVersions();
}

/// <summary>
///     Builder for SupportedProduct.
/// </summary>
public class SupportedProductBuilder : BuilderBase<SupportedProductBuilder, SupportedProduct>, ISupportedProductBuilder
{
    private List<SupportedMajorVersion>           _majorVersions = [];
    private IReadOnlyList<SupportedMajorVersion>? _majorVersionsView;

    public string?                              Name          { get; private set; }
    public string?                              DisplayName   { get; private set; }
    public IReadOnlyList<SupportedMajorVersion> MajorVersions => _majorVersionsView ??= _majorVersions.AsReadOnly();

    public ISupportedProductBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));

        Name = name;

        return this;
    }

    public ISupportedProductBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName = displayName;

        return this;
    }

    public ISupportedProductBuilder AddMajorVersion(SupportedMajorVersion majorVersion)
    {
        ArgumentNullException.ThrowIfNull(majorVersion, nameof(majorVersion));
        EnsureNotBuilt(nameof(AddMajorVersion));

        _majorVersions.Add(majorVersion);

        return this;
    }

    public ISupportedProductBuilder AddMajorVersions(IEnumerable<SupportedMajorVersion> majorVersions)
    {
        ArgumentNullException.ThrowIfNull(majorVersions, nameof(majorVersions));
        EnsureNotBuilt(nameof(AddMajorVersions));

        foreach (var majorVersion in majorVersions)
        {
            if (majorVersion == null)
                throw new ArgumentNullException(nameof(majorVersions), "Major versions collection cannot contain null values.");

            _majorVersions.Add(majorVersion);
        }

        return this;
    }

    public ISupportedProductBuilder RemoveMajorVersion(SupportedMajorVersion majorVersion)
    {
        ArgumentNullException.ThrowIfNull(majorVersion, nameof(majorVersion));
        EnsureNotBuilt(nameof(RemoveMajorVersion));

        _majorVersions.Remove(majorVersion);

        return this;
    }

    public ISupportedProductBuilder RemoveMajorVersions(IEnumerable<SupportedMajorVersion> majorVersions)
    {
        ArgumentNullException.ThrowIfNull(majorVersions, nameof(majorVersions));
        EnsureNotBuilt(nameof(RemoveMajorVersions));

        foreach (var majorVersion in majorVersions)
        {
            if (majorVersion == null)
                throw new ArgumentNullException(nameof(majorVersions), "Major versions collection cannot contain null values.");

            _majorVersions.Remove(majorVersion);
        }

        return this;
    }

    public ISupportedProductBuilder ClearMajorVersions()
    {
        EnsureNotBuilt(nameof(ClearMajorVersions));
        _majorVersions.Clear();

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<SupportedProductBuilder, SupportedProduct>(
            () => (nameof(Name), Name is not null),
            () => (nameof(DisplayName), DisplayName is not null)
        );

    protected override SupportedProduct BuildCore() =>
        new()
        {
            Name          = Name!,
            DisplayName   = DisplayName!,
            MajorVersions = MajorVersions
        };

    protected override void ResetCore()
    {
        Name               = null;
        DisplayName        = null;
        _majorVersions     = [];
        _majorVersionsView = null;
    }
}