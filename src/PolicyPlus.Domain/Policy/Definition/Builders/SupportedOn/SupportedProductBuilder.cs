using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.SupportedOn;

namespace PolicyPlus.Domain.Policy.Definition.Builders.SupportedOn;

/// <summary>
///     Builder interface for SupportedProduct.
/// </summary>
public interface ISupportedProductBuilder : IBuilder<SupportedProduct>
{
    ISupportedProductBuilder WithName(string                       name);
    ISupportedProductBuilder WithDisplayName(string                displayName);
    ISupportedProductBuilder AddMajorVersion(SupportedMajorVersion majorVersion);
}

/// <summary>
///     Builder for SupportedProduct.
/// </summary>
public class SupportedProductBuilder : BuilderBase<SupportedProductBuilder, SupportedProduct>, ISupportedProductBuilder
{
    private readonly List<SupportedMajorVersion> _majorVersions = [];

    public string?                              Name          { get; private set; }
    public string?                              DisplayName   { get; private set; }
    public IReadOnlyList<SupportedMajorVersion> MajorVersions => _majorVersions.AsReadOnly();

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
            MajorVersions = _majorVersions.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Name        = null;
        DisplayName = null;

        _majorVersions.Clear();
    }
}