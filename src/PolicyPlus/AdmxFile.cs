using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;

namespace PolicyPlus;

/// <summary>
///     Represents an ADMX (Administrative Template XML) file that defines policy settings
/// </summary>
public class AdmxFile
{
    private AdmxFile() { }

    /// <summary>
    ///     Loads an ADMX file from the specified path
    /// </summary>
    /// <param name="filePath">Path to the ADMX file</param>
    /// <returns>A populated AdmxFile object</returns>
    public static AdmxFile Load(string filePath)
    {
        // ADMX documentation: https://technet.microsoft.com/en-us/library/cc772138(v=ws.10).aspx
        var admx = new AdmxFile
        {
            SourceFile = filePath
        };

        var xmlDoc = new XmlDocument();
        xmlDoc.Load(filePath);

        var policyDefinitions = xmlDoc.GetElementsByTagName("policyDefinitions")[0];

        foreach (XmlNode child in policyDefinitions.ChildNodes)
        {
            switch (child.LocalName ?? "")
            {
                case "policyNamespaces":
                    ProcessPolicyNamespaces(child, admx);

                    break;

                case "supersededAdm":
                    admx.SupersededAdm = child.Attributes["fileName"]?.Value ?? string.Empty;

                    break;

                case "resources":
                    admx.MinAdmlVersion = decimal.Parse(
                        child.Attributes["minRequiredRevision"]?.Value ?? "0",
                        CultureInfo.InvariantCulture
                    );

                    break;

                case "supportedOn":
                    ProcessSupportedOn(child, admx);

                    break;

                case "categories":
                    ProcessCategories(child, admx);

                    break;

                case "policies":
                    ProcessPolicies(child, admx);

                    break;
            }
        }

        return admx;
    }

    #region Properties

    public string                      AdmxNamespace          { get; private set; } = string.Empty;
    public List<AdmxCategory>          Categories             { get; }              = new();
    public decimal                     MinAdmlVersion         { get; private set; }
    public List<AdmxPolicy>            Policies               { get; }              = new();
    public Dictionary<string, string>  Prefixes               { get; }              = new();
    public List<AdmxProduct>           Products               { get; }              = new();
    public string                      SourceFile             { get; private set; } = string.Empty;
    public string                      SupersededAdm          { get; private set; } = string.Empty;
    public List<AdmxSupportDefinition> SupportedOnDefinitions { get; }              = new();

    #endregion

    #region Helper Methods for Processing ADMX Sections

    /// <summary>
    ///     Process the policy namespaces section of the ADMX file
    /// </summary>
    private static void ProcessPolicyNamespaces(XmlNode policyNamespacesNode, AdmxFile admx)
    {
        foreach (XmlNode policyNamespace in policyNamespacesNode.ChildNodes)
        {
            if (policyNamespace.Attributes?["prefix"]   == null ||
                policyNamespace.Attributes["namespace"] == null)
                continue;

            var prefix      = policyNamespace.Attributes["prefix"].Value;
            var fqNamespace = policyNamespace.Attributes["namespace"].Value;

            if (policyNamespace.LocalName == "target")
                admx.AdmxNamespace = fqNamespace;

            admx.Prefixes.Add(prefix, fqNamespace);
        }
    }

    /// <summary>
    ///     Process the supported products and definitions section of the ADMX file
    /// </summary>
    private static void ProcessSupportedOn(XmlNode supportedOnNode, AdmxFile admx)
    {
        foreach (XmlNode supportInfo in supportedOnNode.ChildNodes)
        {
            if (supportInfo.LocalName == "definitions")
                ProcessSupportDefinitions(supportInfo, admx);
            else if (supportInfo.LocalName == "products")
            {
                // Start the recursive load of product definitions
                LoadProductsRecursively(supportInfo, "product", null, admx);
            }
        }
    }

    /// <summary>
    ///     Process the support definitions section of the ADMX file
    /// </summary>
    private static void ProcessSupportDefinitions(XmlNode definitionsNode, AdmxFile admx)
    {
        foreach (XmlNode supportDef in definitionsNode.ChildNodes)
        {
            if (supportDef.LocalName != "definition")
                continue;

            var definition = new AdmxSupportDefinition
            {
                ID          = supportDef.GetAttributeValue("name"),
                DisplayCode = supportDef.GetAttributeValue("displayName"),
                Logic       = AdmxSupportLogicType.Blank,
                DefinedIn   = admx
            };

            ProcessSupportLogic(supportDef, definition);
            admx.SupportedOnDefinitions.Add(definition);
        }
    }

    /// <summary>
    ///     Process the logical conditions (AND/OR) for support definitions
    /// </summary>
    private static void ProcessSupportLogic(XmlNode supportDef, AdmxSupportDefinition definition)
    {
        foreach (XmlNode logicElement in supportDef.ChildNodes)
        {
            var logicType = logicElement.LocalName switch
            {
                "or"  => AdmxSupportLogicType.AnyOf,
                "and" => AdmxSupportLogicType.AllOf,
                _     => AdmxSupportLogicType.Blank
            };

            if (logicType == AdmxSupportLogicType.Blank)
                continue;

            definition.Logic   = logicType;
            definition.Entries = new List<AdmxSupportEntry>();

            foreach (XmlNode conditionElement in logicElement.ChildNodes)
            {
                if (conditionElement.LocalName == "reference")
                {
                    var product = conditionElement.GetAttributeValue("ref");
                    definition.Entries.Add(new AdmxSupportEntry { ProductID = product, IsRange = false });
                }
                else if (conditionElement.LocalName == "range")
                {
                    var entry = new AdmxSupportEntry { IsRange = true };
                    entry.ProductID = conditionElement.GetAttributeValue("ref");

                    var maxVerAttr = conditionElement.Attributes?["maxVersionIndex"];

                    if (maxVerAttr != null)
                        entry.MaxVersion = int.Parse(maxVerAttr.Value);

                    var minVerAttr = conditionElement.Attributes?["minVersionIndex"];

                    if (minVerAttr != null)
                        entry.MinVersion = int.Parse(minVerAttr.Value);

                    definition.Entries.Add(entry);
                }
            }

            // We found a valid logic element, so we can break
            break;
        }
    }

    /// <summary>
    ///     Process the categories section of the ADMX file
    /// </summary>
    private static void ProcessCategories(XmlNode categoriesNode, AdmxFile admx)
    {
        foreach (XmlNode categoryElement in categoriesNode.ChildNodes)
        {
            if (categoryElement.LocalName != "category")
                continue;

            var category = new AdmxCategory
            {
                ID          = categoryElement.GetAttributeValue("name"),
                DisplayCode = categoryElement.GetAttributeValue("displayName"),
                ExplainCode = categoryElement.GetAttributeOrNull("explainText"),
                DefinedIn   = admx
            };

            if (categoryElement.HasChildNodes)
            {
                var parentCatElement = categoryElement["parentCategory"];

                if (parentCatElement != null)
                    category.ParentID = parentCatElement.GetAttributeValue("ref");
            }

            admx.Categories.Add(category);
        }
    }

    /// <summary>
    ///     Process the policies section of the ADMX file
    /// </summary>
    private static void ProcessPolicies(XmlNode policiesNode, AdmxFile admx)
    {
        var registryValueParser = new PolicyRegistryValueParser();

        foreach (XmlNode polElement in policiesNode.ChildNodes)
        {
            if (polElement.LocalName != "policy")
                continue;

            var policy = CreatePolicyFromXmlNode(polElement, admx, registryValueParser);
            admx.Policies.Add(policy);
        }
    }

    /// <summary>
    ///     Create a policy object from an XML node
    /// </summary>
    private static AdmxPolicy CreatePolicyFromXmlNode(XmlNode polElement, AdmxFile admx, PolicyRegistryValueParser registryValueParser)
    {
        var policy = new AdmxPolicy
        {
            ID              = polElement.GetAttributeValue("name"),
            DefinedIn       = admx,
            DisplayCode     = polElement.GetAttributeValue("displayName"),
            RegistryKey     = polElement.GetAttributeValue("key"),
            ExplainCode     = polElement.GetAttributeOrNull("explainText"),
            PresentationID  = polElement.GetAttributeOrNull("presentation"),
            ClientExtension = polElement.GetAttributeOrNull("clientExtension"),
            RegistryValue   = polElement.GetAttributeOrNull("valueName")
        };

        // Set policy section based on class attribute
        policy.Section = polElement.GetAttributeValue("class") switch
        {
            "Machine" => AdmxPolicySection.Machine,
            "User"    => AdmxPolicySection.User,
            _         => AdmxPolicySection.Both
        };

        // Process affected registry values
        policy.AffectedValues = registryValueParser.LoadOnOffValueList(
            "enabledValue",
            "disabledValue",
            "enabledList",
            "disabledList",
            polElement
        );

        // Process policy child nodes (parentCategory, supportedOn, elements)
        foreach (XmlNode polInfo in polElement.ChildNodes)
        {
            switch (polInfo.LocalName)
            {
                case "parentCategory":
                    policy.CategoryID = polInfo.GetAttributeValue("ref");

                    break;

                case "supportedOn":
                    policy.SupportedCode = polInfo.GetAttributeValue("ref");

                    break;

                case "elements":
                    policy.Elements = ProcessPolicyElements(polInfo, registryValueParser);

                    break;
            }
        }

        return policy;
    }

    /// <summary>
    ///     Process policy UI elements
    /// </summary>
    private static List<PolicyElement> ProcessPolicyElements(XmlNode elementsNode, PolicyRegistryValueParser registryValueParser)
    {
        var elements = new List<PolicyElement>();

        foreach (XmlNode uiElement in elementsNode.ChildNodes)
        {
            PolicyElement entry = uiElement.LocalName switch
            {
                "decimal"   => CreateDecimalElement(uiElement),
                "boolean"   => CreateBooleanElement(uiElement, registryValueParser),
                "text"      => CreateTextElement(uiElement),
                "list"      => CreateListElement(uiElement),
                "enum"      => CreateEnumElement(uiElement, registryValueParser),
                "multiText" => new MultiTextPolicyElement(),
                _           => null
            };

            if (entry == null)
                continue;

            entry.ClientExtension = uiElement.GetAttributeOrNull("clientExtension");
            entry.RegistryKey     = uiElement.GetAttributeOrNull("key");

            if (string.IsNullOrEmpty(entry.RegistryValue))
                entry.RegistryValue = uiElement.GetAttributeOrNull("valueName");

            entry.ID          = uiElement.GetAttributeValue("id");
            entry.ElementType = uiElement.LocalName;

            elements.Add(entry);
        }

        return elements;
    }

    /// <summary>
    ///     Create a decimal policy element from an XML node
    /// </summary>
    private static DecimalPolicyElement CreateDecimalElement(XmlNode uiElement) =>
        new()
        {
            Minimum     = Convert.ToUInt32(uiElement.GetAttributeOrDefault("minValue",     "0")),
            Maximum     = Convert.ToUInt32(uiElement.GetAttributeOrDefault("maxValue",     uint.MaxValue.ToString())),
            NoOverwrite = Convert.ToBoolean(uiElement.GetAttributeOrDefault("soft",        "false")),
            StoreAsText = Convert.ToBoolean(uiElement.GetAttributeOrDefault("storeAsText", "false"))
        };

    /// <summary>
    ///     Create a boolean policy element from an XML node
    /// </summary>
    private static BooleanPolicyElement CreateBooleanElement(XmlNode uiElement, PolicyRegistryValueParser registryValueParser)
    {
        var boolEntry = new BooleanPolicyElement
        {
            AffectedRegistry = registryValueParser.LoadOnOffValueList("trueValue", "falseValue", "trueList", "falseList", uiElement)
        };

        return boolEntry;
    }

    /// <summary>
    ///     Create a text policy element from an XML node
    /// </summary>
    private static TextPolicyElement CreateTextElement(XmlNode uiElement) =>
        new()
        {
            MaxLength   = Convert.ToInt32(uiElement.GetAttributeOrDefault("maxLength",    "255")),
            Required    = Convert.ToBoolean(uiElement.GetAttributeOrDefault("required",   "false")),
            RegExpandSz = Convert.ToBoolean(uiElement.GetAttributeOrDefault("expandable", "false")),
            NoOverwrite = Convert.ToBoolean(uiElement.GetAttributeOrDefault("soft",       "false"))
        };

    /// <summary>
    ///     Create a list policy element from an XML node
    /// </summary>
    private static ListPolicyElement CreateListElement(XmlNode uiElement)
    {
        var listEntry = new ListPolicyElement
        {
            NoPurgeOthers     = Convert.ToBoolean(uiElement.GetAttributeOrDefault("additive",      "false")),
            RegExpandSz       = Convert.ToBoolean(uiElement.GetAttributeOrDefault("expandable",    "false")),
            UserProvidesNames = Convert.ToBoolean(uiElement.GetAttributeOrDefault("explicitValue", "false")),
            HasPrefix         = uiElement.Attributes?["valuePrefix"] != null,
            RegistryValue     = uiElement.GetAttributeOrNull("valuePrefix")
        };

        return listEntry;
    }

    /// <summary>
    ///     Create an enum policy element from an XML node
    /// </summary>
    private static EnumPolicyElement CreateEnumElement(XmlNode uiElement, PolicyRegistryValueParser registryValueParser)
    {
        var enumEntry = new EnumPolicyElement
        {
            Required = Convert.ToBoolean(uiElement.GetAttributeOrDefault("required", "false")),
            Items    = new List<EnumPolicyElementItem>()
        };

        foreach (XmlNode itemElement in uiElement.ChildNodes)
        {
            if (itemElement.LocalName != "item")
                continue;

            var enumItem = new EnumPolicyElementItem
            {
                DisplayCode = itemElement.GetAttributeValue("displayName")
            };

            foreach (XmlNode valElement in itemElement.ChildNodes)
            {
                if (valElement.LocalName == "value")
                    enumItem.Value = registryValueParser.LoadRegistryValue(valElement);
                else if (valElement.LocalName == "valueList")
                    enumItem.ValueList = registryValueParser.LoadSingleRegistryList(valElement);
            }

            enumEntry.Items.Add(enumItem);
        }

        return enumEntry;
    }

    /// <summary>
    ///     Recursively load product definitions
    /// </summary>
    private static void LoadProductsRecursively(XmlNode node, string childTagName, AdmxProduct parent, AdmxFile admx)
    {
        foreach (XmlNode subproductElement in node.ChildNodes)
        {
            if ((subproductElement.LocalName ?? "") != (childTagName ?? ""))
                continue;

            var product = new AdmxProduct
            {
                ID          = subproductElement.GetAttributeValue("name"),
                DisplayCode = subproductElement.GetAttributeValue("displayName"),
                Parent      = parent,
                DefinedIn   = admx
            };

            if (parent != null)
                product.Version = Convert.ToInt32(subproductElement.GetAttributeValue("versionIndex"));

            admx.Products.Add(product);

            if (parent == null)
            {
                product.Type = AdmxProductType.Product;
                LoadProductsRecursively(subproductElement, "majorVersion", product, admx);
            }
            else if (parent.Parent == null)
            {
                product.Type = AdmxProductType.MajorRevision;
                LoadProductsRecursively(subproductElement, "minorVersion", product, admx);
            }
            else
                product.Type = AdmxProductType.MinorRevision;
        }
    }

    #endregion
}

/// <summary>
///     Helper class for parsing registry values from XML nodes
/// </summary>
internal class PolicyRegistryValueParser
{
    /// <summary>
    ///     Load a registry value from an XML node
    /// </summary>
    public PolicyRegistryValue LoadRegistryValue(XmlNode node)
    {
        var regItem = new PolicyRegistryValue();

        foreach (XmlNode subElement in node.ChildNodes)
        {
            if (subElement.LocalName == "delete")
            {
                regItem.RegistryType = PolicyRegistryValueType.Delete;

                break;
            }

            if (subElement.LocalName == "decimal")
            {
                regItem.RegistryType = PolicyRegistryValueType.Numeric;
                regItem.NumberValue  = Convert.ToUInt32(subElement.GetAttributeValue("value"));

                break;
            }

            if (subElement.LocalName != "string")
                continue;

            regItem.RegistryType = PolicyRegistryValueType.Text;
            regItem.StringValue  = subElement.InnerText;

            break;
        }

        return regItem;
    }

    /// <summary>
    ///     Load a registry list from an XML node
    /// </summary>
    public PolicyRegistrySingleList LoadSingleRegistryList(XmlNode node)
    {
        var singleList = new PolicyRegistrySingleList
        {
            DefaultRegistryKey = node.GetAttributeOrNull("defaultKey"),
            AffectedValues     = new List<PolicyRegistryListEntry>()
        };

        foreach (XmlNode itemElement in node.ChildNodes)
        {
            if (itemElement.LocalName != "item")
                continue;

            var listEntry = new PolicyRegistryListEntry
            {
                RegistryValue = itemElement.GetAttributeValue("valueName"),
                RegistryKey   = itemElement.GetAttributeOrNull("key")
            };

            foreach (XmlNode valElement in itemElement.ChildNodes)
            {
                if (valElement.LocalName != "value")
                    continue;

                listEntry.Value = LoadRegistryValue(valElement);

                break;
            }

            singleList.AffectedValues.Add(listEntry);
        }

        return singleList;
    }

    /// <summary>
    ///     Load registry values for on/off settings
    /// </summary>
    public PolicyRegistryList LoadOnOffValueList(
        string  onValueName,
        string  offValueName,
        string  onListName,
        string  offListName,
        XmlNode node
    )
    {
        var regList = new PolicyRegistryList();

        foreach (XmlNode subElement in node.ChildNodes)
        {
            if (string.Equals(subElement.Name ?? "", onValueName ?? "", StringComparison.OrdinalIgnoreCase))
                regList.OnValue = LoadRegistryValue(subElement);

            else if (string.Equals(subElement.Name ?? "", offValueName ?? "", StringComparison.OrdinalIgnoreCase))
                regList.OffValue = LoadRegistryValue(subElement);

            else if (string.Equals(subElement.Name ?? "", onListName ?? "", StringComparison.OrdinalIgnoreCase))
                regList.OnValueList = LoadSingleRegistryList(subElement);

            else if (string.Equals(subElement.Name ?? "", offListName ?? "", StringComparison.OrdinalIgnoreCase))
                regList.OffValueList = LoadSingleRegistryList(subElement);
        }

        return regList;
    }
}

/// <summary>
///     Extension methods for XmlNode to simplify attribute access
/// </summary>
internal static class XmlNodeExtensions
{
    /// <summary>
    ///     Get attribute value or throw if not found
    /// </summary>
    public static string GetAttributeValue(this XmlNode node, string attributeName)
    {
        var attribute = node.Attributes?[attributeName];

        if (attribute == null)
            throw new ArgumentException($"Required attribute '{attributeName}' not found on node '{node.Name}'");

        return attribute.Value;
    }

    /// <summary>
    ///     Get attribute value or return null if not found
    /// </summary>
    public static string GetAttributeOrNull(this XmlNode node, string attributeName) => node.Attributes?[attributeName]?.Value;

    /// <summary>
    ///     Get attribute value or return default if not found
    /// </summary>
    public static string GetAttributeOrDefault(this XmlNode node, string attributeName, string defaultValue) =>
        node.Attributes?[attributeName]?.Value ?? defaultValue;
}