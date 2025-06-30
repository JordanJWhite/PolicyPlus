using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.References;

namespace PolicyPlus.Domain.Policy.Definition.Builders.References;

/// <summary>
///     Builder interface for FileReference.
/// </summary>
public interface IFileReferenceBuilder : IBuilder<FileReference>
{
    IFileReferenceBuilder WithFileName(string fileName);
}

/// <summary>
///     Builder for FileReference.
/// </summary>
public class FileReferenceBuilder : BuilderBase<FileReferenceBuilder, FileReference>, IFileReferenceBuilder
{
    public string? FileName { get; private set; }

    public IFileReferenceBuilder WithFileName(string fileName)
    {
        ArgumentNullException.ThrowIfNull(fileName, nameof(fileName));
        EnsureNotBuilt(nameof(WithFileName));

        FileName = fileName;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<FileReferenceBuilder, FileReference>(() => (nameof(FileName),
                                                                                                            FileName is not null)
        );

    protected override FileReference BuildCore() =>
        new()
        {
            FileName = FileName!
        };

    protected override void ResetCore() => FileName = null;
}