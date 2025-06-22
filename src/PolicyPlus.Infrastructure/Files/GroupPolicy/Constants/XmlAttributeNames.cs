namespace PolicyPlus.Infrastructure.Files.GroupPolicy.Constants;

/// <summary>
///     Contains XML attribute names for ADMX/ADML parsing.
/// </summary>
internal static class XmlAttributeNames
{
    // Common attributes
    public const string Id            = "id";
    public const string Name          = "name";
    public const string DisplayName   = "displayName";
    public const string Revision      = "revision";
    public const string SchemaVersion = "schemaVersion";
    public const string Key           = "key";
    public const string ValueName     = "valueName";
    public const string Ref           = "ref";
    public const string RefId         = "refId";

    // Namespace attributes
    public const string Prefix    = "prefix";
    public const string Namespace = "namespace";

    // File reference
    public const string FileName = "fileName";

    // Localization
    public const string MinRequiredRevision = "minRequiredRevision";
    public const string FallbackCulture     = "fallbackCulture";

    // Policy attributes
    public const string Class           = "class";
    public const string ExplainText     = "explainText";
    public const string Presentation    = "presentation";
    public const string ClientExtension = "clientExtension";

    // Element attributes
    public const string Required      = "required";
    public const string MinValue      = "minValue";
    public const string MaxValue      = "maxValue";
    public const string StoreAsText   = "storeAsText";
    public const string Soft          = "soft";
    public const string MaxLength     = "maxLength";
    public const string Expandable    = "expandable";
    public const string ValuePrefix   = "valuePrefix";
    public const string Additive      = "additive";
    public const string ExplicitValue = "explicitValue";
    public const string MaxStrings    = "maxStrings";

    // SupportedOn attributes
    public const string VersionIndex    = "versionIndex";
    public const string MinVersionIndex = "minVersionIndex";
    public const string MaxVersionIndex = "maxVersionIndex";

    // Annotation attributes
    public const string Application = "application";

    // Value attributes
    public const string Value      = "value";
    public const string DefaultKey = "defaultKey";

    // Presentation attributes
    public const string DefaultChecked = "defaultChecked";
    public const string Spin           = "spin";
    public const string SpinStep       = "spinStep";
    public const string NoSort         = "noSort";
    public const string DefaultItem    = "defaultItem";
    public const string ShowAsDialog   = "showAsDialog";
    public const string DefaultHeight  = "defaultHeight";
}