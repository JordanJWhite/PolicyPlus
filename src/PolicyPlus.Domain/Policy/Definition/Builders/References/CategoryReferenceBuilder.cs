using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.References;

namespace PolicyPlus.Domain.Policy.Definition.Builders.References;

/// <summary>
///     Builder interface for CategoryReference.
/// </summary>
public interface ICategoryReferenceBuilder : IBuilder<CategoryReference>
{
    ICategoryReferenceBuilder WithRef(string reference);
}

/// <summary>
///     Builder for CategoryReference.
/// </summary>
public class CategoryReferenceBuilder : BuilderBase<CategoryReferenceBuilder, CategoryReference>, ICategoryReferenceBuilder
{
    public string? Ref { get; private set; }

    public ICategoryReferenceBuilder WithRef(string reference)
    {
        ArgumentNullException.ThrowIfNull(reference, nameof(reference));
        EnsureNotBuilt(nameof(WithRef));

        Ref = reference;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper
           .ThrowIfRequiredPropertiesMissing<CategoryReferenceBuilder, CategoryReference>(() => (nameof(Ref), Ref is not null));

    protected override CategoryReference BuildCore() =>
        new()
        {
            Ref = Ref!
        };

    protected override void ResetCore() => Ref = null;
}