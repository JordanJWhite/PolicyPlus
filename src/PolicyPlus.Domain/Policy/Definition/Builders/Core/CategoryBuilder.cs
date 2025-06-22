using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Builders.Helpers;
using PolicyPlus.Domain.Policy.Definition.Model.Core;
using PolicyPlus.Domain.Policy.Definition.Model.References;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Core;

/// <summary>
///     Builder interface for Category.
/// </summary>
public interface ICategoryBuilder : IBuilder<Category>
{
    ICategoryBuilder WithName(string                       name);
    ICategoryBuilder WithDisplayName(string                displayName);
    ICategoryBuilder WithExplainText(string?               explainText);
    ICategoryBuilder WithParentCategory(CategoryReference? parentCategory);
    ICategoryBuilder WithKeywords(string?                  keywords);
    ICategoryBuilder AddAnnotation(Annotation              annotation);
    ICategoryBuilder AddSeeAlso(string                     seeAlso);
}

/// <summary>
///     Builder for Category.
/// </summary>
public class CategoryBuilder : BuilderBase<CategoryBuilder, Category>, ICategoryBuilder
{
    // Collections to hold annotations and "see also" references.
    private List<Annotation> _annotations = [];
    private List<string>     _seeAlso     = [];

    // Views for read-only access to the collections.
    private IReadOnlyList<Annotation>? _annotationsView;
    private IReadOnlyList<string>?     _seeAlsoView;

    public string?                   Name           { get; private set; }
    public string?                   DisplayName    { get; private set; }
    public string?                   ExplainText    { get; private set; }
    public CategoryReference?        ParentCategory { get; private set; }
    public string?                   Keywords       { get; private set; }
    public IReadOnlyList<Annotation> Annotations    => _annotationsView ??= _annotations.AsReadOnly();
    public IReadOnlyList<string>     SeeAlso        => _seeAlsoView ??= _seeAlso.AsReadOnly();

    public ICategoryBuilder WithName(string name)
    {
        ArgumentNullException.ThrowIfNull(name, nameof(name));
        EnsureNotBuilt(nameof(WithName));

        Name = name;

        return this;
    }

    public ICategoryBuilder WithDisplayName(string displayName)
    {
        ArgumentNullException.ThrowIfNull(displayName, nameof(displayName));
        EnsureNotBuilt(nameof(WithDisplayName));

        DisplayName = displayName;

        return this;
    }

    public ICategoryBuilder WithExplainText(string? explainText)
    {
        EnsureNotBuilt(nameof(WithExplainText));
        ExplainText = explainText;

        return this;
    }

    public ICategoryBuilder WithParentCategory(CategoryReference? parentCategory)
    {
        EnsureNotBuilt(nameof(WithParentCategory));
        ParentCategory = parentCategory;

        return this;
    }

    public ICategoryBuilder WithKeywords(string? keywords)
    {
        EnsureNotBuilt(nameof(WithKeywords));
        Keywords = keywords;

        return this;
    }

    public ICategoryBuilder AddAnnotation(Annotation annotation)
    {
        ArgumentNullException.ThrowIfNull(annotation, nameof(annotation));
        EnsureNotBuilt(nameof(AddAnnotation));

        _annotations.Add(annotation);

        return this;
    }

    public ICategoryBuilder AddSeeAlso(string seeAlso)
    {
        ArgumentNullException.ThrowIfNull(seeAlso, nameof(seeAlso));
        EnsureNotBuilt(nameof(AddSeeAlso));

        _seeAlso.Add(seeAlso);

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<CategoryBuilder, Category>(
            () => (nameof(Name), Name is not null),
            () => (nameof(DisplayName), DisplayName is not null)
        );

    protected override Category BuildCore() =>
        new()
        {
            Name           = Name!,
            DisplayName    = DisplayName!,
            ExplainText    = ExplainText,
            ParentCategory = ParentCategory,
            Keywords       = Keywords,
            Annotations    = _annotations.AsReadOnly(),
            SeeAlso        = _seeAlso.AsReadOnly()
        };

    protected override void ResetCore()
    {
        Name             = null;
        DisplayName      = null;
        ExplainText      = null;
        ParentCategory   = null;
        Keywords         = null;
        _annotations     = [];
        _seeAlso         = [];
        _annotationsView = null;
        _seeAlsoView     = null;
    }
}