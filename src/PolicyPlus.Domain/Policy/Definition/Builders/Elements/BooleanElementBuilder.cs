using PolicyPlus.Domain.Builders.Base;
using PolicyPlus.Domain.Policy.Definition.Model.Elements;
using PolicyPlus.Domain.Policy.Definition.Model.Values;

namespace PolicyPlus.Domain.Policy.Definition.Builders.Elements;

/// <summary>
///     Builder interface for BooleanElement.
/// </summary>
public interface IBooleanElementBuilder : IBuilder<BooleanElement>
{
    IBooleanElementBuilder WithId(string               id);
    IBooleanElementBuilder WithClientExtension(string? clientExtension);
    IBooleanElementBuilder WithKey(string?             key);
    IBooleanElementBuilder WithValueName(string?       valueName);
    IBooleanElementBuilder WithTrueValue(IValue?       trueValue);
    IBooleanElementBuilder WithFalseValue(IValue?      falseValue);
    IBooleanElementBuilder WithTrueList(ValueList?     trueList);
    IBooleanElementBuilder WithFalseList(ValueList?    falseList);
}

/// <summary>
///     Builder for BooleanElement.
/// </summary>
public class BooleanElementBuilder : BuilderBase<BooleanElementBuilder, BooleanElement>, IBooleanElementBuilder
{
    public string?    Id              { get; private set; }
    public string?    ClientExtension { get; private set; }
    public string?    Key             { get; private set; }
    public string?    ValueName       { get; private set; }
    public IValue?    TrueValue       { get; private set; }
    public IValue?    FalseValue      { get; private set; }
    public ValueList? TrueList        { get; private set; }
    public ValueList? FalseList       { get; private set; }

    public IBooleanElementBuilder WithId(string id)
    {
        ArgumentNullException.ThrowIfNull(id, nameof(id));
        EnsureNotBuilt(nameof(WithId));
        Id = id;

        return this;
    }

    public IBooleanElementBuilder WithClientExtension(string? clientExtension)
    {
        EnsureNotBuilt(nameof(WithClientExtension));
        ClientExtension = clientExtension;

        return this;
    }

    public IBooleanElementBuilder WithKey(string? key)
    {
        EnsureNotBuilt(nameof(WithKey));
        Key = key;

        return this;
    }

    public IBooleanElementBuilder WithValueName(string? valueName)
    {
        EnsureNotBuilt(nameof(WithValueName));
        ValueName = valueName;

        return this;
    }

    public IBooleanElementBuilder WithTrueValue(IValue? trueValue)
    {
        EnsureNotBuilt(nameof(WithTrueValue));
        TrueValue = trueValue;

        return this;
    }

    public IBooleanElementBuilder WithFalseValue(IValue? falseValue)
    {
        EnsureNotBuilt(nameof(WithFalseValue));
        FalseValue = falseValue;

        return this;
    }

    public IBooleanElementBuilder WithTrueList(ValueList? trueList)
    {
        EnsureNotBuilt(nameof(WithTrueList));
        TrueList = trueList;

        return this;
    }

    public IBooleanElementBuilder WithFalseList(ValueList? falseList)
    {
        EnsureNotBuilt(nameof(WithFalseList));
        FalseList = falseList;

        return this;
    }

    protected override void ValidateRequiredProperties() =>
        BuilderExceptionHelper.ThrowIfRequiredPropertiesMissing<BooleanElementBuilder, BooleanElement>(() => (nameof(Id), Id is not null));

    protected override BooleanElement BuildCore() =>
        new()
        {
            Id              = Id!,
            ClientExtension = ClientExtension,
            Key             = Key,
            ValueName       = ValueName,
            TrueValue       = TrueValue,
            FalseValue      = FalseValue,
            TrueList        = TrueList,
            FalseList       = FalseList
        };

    protected override void ResetCore()
    {
        Id              = null;
        ClientExtension = null;
        Key             = null;
        ValueName       = null;
        TrueValue       = null;
        FalseValue      = null;
        TrueList        = null;
        FalseList       = null;
    }
}