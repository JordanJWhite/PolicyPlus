using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Presentation.Elements;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Presentation.Elements;

/// <summary>
///     Builder interface for ComboBoxElement.
/// </summary>
public interface IComboBoxElementBuilder : IBuilder<ComboBoxElement>
{
    IComboBoxElementBuilder WithLabel(string                      label);
    IComboBoxElementBuilder WithDefaultValue(string?              defaultValue);
    IComboBoxElementBuilder AddSuggestion(string                  suggestion);
    IComboBoxElementBuilder AddSuggestions(IEnumerable<string>    suggestions);
    IComboBoxElementBuilder RemoveSuggestion(string               suggestion);
    IComboBoxElementBuilder RemoveSuggestions(IEnumerable<string> suggestions);
    IComboBoxElementBuilder ClearSuggestions();
    IComboBoxElementBuilder WithNoSort(bool  noSort);
    IComboBoxElementBuilder WithRefId(string refId);
}

/// <summary>
///     Builder for ComboBoxElement.
/// </summary>
public class ComboBoxElementBuilder : BuilderBase<ComboBoxElementBuilder, ComboBoxElement>, IComboBoxElementBuilder
{
    private List<string>           _suggestions = [];
    private IReadOnlyList<string>? _suggestionsView;

    public string?               Label        { get; private set; }
    public string?               DefaultValue { get; private set; }
    public IReadOnlyList<string> Suggestions  => _suggestionsView ??= _suggestions.AsReadOnly();
    public bool                  NoSort       { get; private set; }
    public string?               RefId        { get; private set; }

    public IComboBoxElementBuilder WithLabel(string label)
    {
        ArgumentNullException.ThrowIfNull(label, nameof(label));
        EnsureNotBuilt(nameof(WithLabel));

        Label = label;

        return this;
    }

    public IComboBoxElementBuilder WithDefaultValue(string? defaultValue)
    {
        EnsureNotBuilt(nameof(WithDefaultValue));
        DefaultValue = defaultValue;

        return this;
    }

    public IComboBoxElementBuilder AddSuggestion(string suggestion)
    {
        ArgumentNullException.ThrowIfNull(suggestion, nameof(suggestion));
        EnsureNotBuilt(nameof(AddSuggestion));

        _suggestions.Add(suggestion);

        return this;
    }

    public IComboBoxElementBuilder AddSuggestions(IEnumerable<string> suggestions)
    {
        ArgumentNullException.ThrowIfNull(suggestions, nameof(suggestions));
        EnsureNotBuilt(nameof(AddSuggestions));

        foreach (var suggestion in suggestions)
        {
            if (suggestion == null)
                throw new ArgumentNullException(nameof(suggestions), "Suggestions collection cannot contain null values.");

            _suggestions.Add(suggestion);
        }

        return this;
    }

    public IComboBoxElementBuilder RemoveSuggestion(string suggestion)
    {
        ArgumentNullException.ThrowIfNull(suggestion, nameof(suggestion));
        EnsureNotBuilt(nameof(RemoveSuggestion));

        _suggestions.Remove(suggestion);

        return this;
    }

    public IComboBoxElementBuilder RemoveSuggestions(IEnumerable<string> suggestions)
    {
        ArgumentNullException.ThrowIfNull(suggestions, nameof(suggestions));
        EnsureNotBuilt(nameof(RemoveSuggestions));

        foreach (var suggestion in suggestions)
        {
            if (suggestion == null)
                throw new ArgumentNullException(nameof(suggestions), "Suggestions collection cannot contain null values.");

            _suggestions.Remove(suggestion);
        }

        return this;
    }

    public IComboBoxElementBuilder ClearSuggestions()
    {
        EnsureNotBuilt(nameof(ClearSuggestions));
        _suggestions.Clear();
        _suggestionsView = null;

        return this;
    }

    public IComboBoxElementBuilder WithNoSort(bool noSort)
    {
        EnsureNotBuilt(nameof(WithNoSort));
        NoSort = noSort;

        return this;
    }

    public IComboBoxElementBuilder WithRefId(string refId)
    {
        ArgumentNullException.ThrowIfNull(refId, nameof(refId));
        EnsureNotBuilt(nameof(WithRefId));

        RefId = refId;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<ComboBoxElementBuilder, ComboBoxElement>(
            () => (nameof(Label), Label is not null),
            () => (nameof(RefId), RefId is not null)
        );

    protected override ComboBoxElement BuildCore() =>
        new()
        {
            Label        = Label!,
            DefaultValue = DefaultValue,
            Suggestions  = Suggestions,
            NoSort       = NoSort,
            RefId        = RefId!
        };

    protected override void ResetCore()
    {
        Label            = null;
        DefaultValue     = null;
        NoSort           = false;
        RefId            = null;
        _suggestions     = [];
        _suggestionsView = null;
    }
}