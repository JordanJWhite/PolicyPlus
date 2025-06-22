namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Constants;

/// <summary>
///     Contains XML element names for ADMX/ADML parsing.
/// </summary>
internal static class XmlElementNames
{
    // Root elements
    public const string PolicyDefinitions         = "policyDefinitions";
    public const string PolicyDefinitionResources = "policyDefinitionResources";

    // Common elements
    public const string Annotation   = "annotation";
    public const string DisplayName  = "displayName";
    public const string Description  = "description";
    public const string String       = "string";
    public const string Presentation = "presentation";

    // PolicyDefinitions elements
    public const string PolicyNamespaces = "policyNamespaces";
    public const string Target           = "target";
    public const string Using            = "using";
    public const string SupersededAdm    = "supersededAdm";
    public const string Resources        = "resources";
    public const string SupportedOn      = "supportedOn";
    public const string Categories       = "categories";
    public const string Category         = "category";
    public const string Policies         = "policies";
    public const string Policy           = "policy";
    public const string ParentCategory   = "parentCategory";
    public const string SeeAlso          = "seeAlso";
    public const string Keywords         = "keywords";
    public const string EnabledValue     = "enabledValue";
    public const string DisabledValue    = "disabledValue";
    public const string EnabledList      = "enabledList";
    public const string DisabledList     = "disabledList";
    public const string Elements         = "elements";

    // Policy element types
    public const string Boolean     = "boolean";
    public const string Decimal     = "decimal";
    public const string Text        = "text";
    public const string Enum        = "enum";
    public const string List        = "list";
    public const string LongDecimal = "longDecimal";
    public const string MultiText   = "multiText";

    // Boolean element specific
    public const string TrueValue  = "trueValue";
    public const string FalseValue = "falseValue";
    public const string TrueList   = "trueList";
    public const string FalseList  = "falseList";

    // Enumeration element specific
    public const string Item      = "item";
    public const string Value     = "value";
    public const string ValueList = "valueList";

    // Value types
    public const string Delete           = "delete";
    public const string StringValue      = "string";
    public const string DecimalValue     = "decimal";
    public const string LongDecimalValue = "longDecimal";

    // SupportedOn elements
    public const string Products     = "products";
    public const string Product      = "product";
    public const string Definitions  = "definitions";
    public const string Definition   = "definition";
    public const string MajorVersion = "majorVersion";
    public const string MinorVersion = "minorVersion";
    public const string Or           = "or";
    public const string And          = "and";
    public const string Range        = "range";
    public const string Reference    = "reference";

    // Presentation elements
    public const string StringTable        = "stringTable";
    public const string PresentationTable  = "presentationTable";
    public const string DecimalTextBox     = "decimalTextBox";
    public const string LongDecimalTextBox = "longDecimalTextBox";
    public const string TextBox            = "textBox";
    public const string CheckBox           = "checkBox";
    public const string ComboBox           = "comboBox";
    public const string DropdownList       = "dropdownList";
    public const string ListBox            = "listBox";
    public const string MultiTextBox       = "multiTextBox";
    public const string Label              = "label";
    public const string DefaultValue       = "defaultValue";
    public const string Default            = "default";
    public const string Suggestion         = "suggestion";
}