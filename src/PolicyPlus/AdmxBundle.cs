using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace PolicyPlus;

public class AdmxBundle
{
    private readonly Dictionary<string, AdmxFile> _namespaces = [];

    // Temporary lists from ADMX files that haven't been integrated yet
    private readonly List<AdmxCategory>          _rawCategories = [];
    private readonly List<AdmxPolicy>            _rawPolicies   = [];
    private readonly List<AdmxProduct>           _rawProducts   = [];
    private readonly List<AdmxSupportDefinition> _rawSupport    = [];

    private readonly Dictionary<AdmxFile, AdmlFile> _sourceFiles = [];

    // Lists of top-level items only
    public Dictionary<string, PolicyPlusCategory> Categories = [];

    // Lists that include all items, even those that are children of others
    public Dictionary<string, PolicyPlusCategory> FlatCategories = [];
    public Dictionary<string, PolicyPlusProduct>  FlatProducts   = [];
    public Dictionary<string, PolicyPlusPolicy>   Policies       = [];

    public Dictionary<string, PolicyPlusProduct> Products           = [];
    public Dictionary<string, PolicyPlusSupport> SupportDefinitions = [];

    public IReadOnlyDictionary<AdmxFile, AdmlFile> Sources => _sourceFiles;

    public IEnumerable<AdmxLoadFailure> LoadFolder(string path, string languageCode)
    {
        var fails = new List<AdmxLoadFailure>();

        foreach (var file in Directory.EnumerateFiles(path))
        {
            if (file.ToLowerInvariant()
                    .EndsWith(".admx"))
            {
                var fail = AddSingleAdmx(file, languageCode);

                if (fail is not null)
                    fails.Add(fail);
            }
        }

        BuildStructures();

        return fails;
    }

    public IEnumerable<AdmxLoadFailure> LoadFile(string path, string languageCode)
    {
        var fail = AddSingleAdmx(path, languageCode);
        BuildStructures();

        return fail is null ? Array.Empty<AdmxLoadFailure>() : new[] { fail };
    }

    private AdmxLoadFailure AddSingleAdmx(string admxPath, string languageCode)
    {
        // Load ADMX file
        AdmxFile admx;
        AdmlFile adml;

        try { admx = AdmxFile.Load(admxPath); }
        catch (XmlException ex) { return new AdmxLoadFailure(AdmxLoadFailType.BadAdmxParse, admxPath, ex.Message); }
        catch (Exception ex) { return new AdmxLoadFailure(AdmxLoadFailType.BadAdmx,         admxPath, ex.Message); }

        if (_namespaces.ContainsKey(admx.AdmxNamespace))
            return new AdmxLoadFailure(AdmxLoadFailType.DuplicateNamespace, admxPath, admx.AdmxNamespace);

        // Find the ADML file
        var fileTitle = Path.GetFileName(admxPath);
        var admlPath  = Path.ChangeExtension(admxPath.Replace(fileTitle, languageCode + @"\" + fileTitle), "adml");

        if (!File.Exists(admlPath))
        {
            var language = languageCode.Split('-')[0];

            foreach (var langSubdir in Directory.EnumerateDirectories(Path.GetDirectoryName(admxPath)))
            {
                var langSubdirTitle = Path.GetFileName(langSubdir);
                var subDirLanguage  = langSubdirTitle.Split('-')[0];

                if (subDirLanguage != language)
                    continue;

                var similarLanguagePath = Path.ChangeExtension(admxPath.Replace(fileTitle, langSubdirTitle + @"\" + fileTitle), "adml");

                if (File.Exists(similarLanguagePath))
                {
                    admlPath = similarLanguagePath;

                    break;
                }
            }
        }

        if (!File.Exists(admlPath))
            admlPath = Path.ChangeExtension(admxPath.Replace(fileTitle, @"en-US\" + fileTitle), "adml");

        if (!File.Exists(admlPath))
            return new AdmxLoadFailure(AdmxLoadFailType.NoAdml, admxPath);

        // Load the ADML
        try { adml = AdmlFile.Load(admlPath); }
        catch (XmlException ex) { return new AdmxLoadFailure(AdmxLoadFailType.BadAdmlParse, admxPath, ex.Message); }
        catch (Exception ex) { return new AdmxLoadFailure(AdmxLoadFailType.BadAdml,         admxPath, ex.Message); }

        // Stage the raw ADMX info for BuildStructures
        _rawCategories.AddRange(admx.Categories);
        _rawProducts.AddRange(admx.Products);
        _rawPolicies.AddRange(admx.Policies);
        _rawSupport.AddRange(admx.SupportedOnDefinitions);
        _sourceFiles.Add(admx, adml);
        _namespaces.Add(admx.AdmxNamespace, admx);

        return null;
    }

    private void BuildStructures()
    {
        var catIds     = new Dictionary<string, PolicyPlusCategory>();
        var productIds = new Dictionary<string, PolicyPlusProduct>();
        var supIds     = new Dictionary<string, PolicyPlusSupport>();
        var polIds     = new Dictionary<string, PolicyPlusPolicy>();
        PolicyPlusCategory FindCatById(string uid) => FindInTempOrFlat(uid, catIds, FlatCategories);
        PolicyPlusSupport  FindSupById(string uid) => FindInTempOrFlat(uid, supIds, SupportDefinitions);

        PolicyPlusProduct FindProductById(string uid) => FindInTempOrFlat(uid, productIds, FlatProducts);

        // First pass: Build the structures without resolving references
        foreach (var rawCat in _rawCategories)
        {
            var cat = new PolicyPlusCategory();
            cat.DisplayName        = ResolveString(rawCat.DisplayCode, rawCat.DefinedIn);
            cat.DisplayExplanation = ResolveString(rawCat.ExplainCode, rawCat.DefinedIn);
            cat.UniqueID           = QualifyName(rawCat.ID, rawCat.DefinedIn);
            cat.RawCategory        = rawCat;
            catIds.Add(cat.UniqueID, cat);
        }

        foreach (var rawProduct in _rawProducts)
        {
            var product = new PolicyPlusProduct();
            product.DisplayName = ResolveString(rawProduct.DisplayCode, rawProduct.DefinedIn);
            product.UniqueID    = QualifyName(rawProduct.ID, rawProduct.DefinedIn);
            product.RawProduct  = rawProduct;
            productIds.Add(product.UniqueID, product);
        }

        foreach (var rawSup in _rawSupport)
        {
            var sup = new PolicyPlusSupport();
            sup.DisplayName = ResolveString(rawSup.DisplayCode, rawSup.DefinedIn);
            sup.UniqueID    = QualifyName(rawSup.ID, rawSup.DefinedIn);

            if (rawSup.Entries is not null)
            {
                foreach (var rawSupEntry in rawSup.Entries)
                {
                    var supEntry = new PolicyPlusSupportEntry();
                    supEntry.RawSupportEntry = rawSupEntry;
                    sup.Elements.Add(supEntry);
                }
            }

            sup.RawSupport = rawSup;
            supIds.Add(sup.UniqueID, sup);
        }

        foreach (var rawPol in _rawPolicies)
        {
            var pol = new PolicyPlusPolicy();
            pol.DisplayExplanation = ResolveString(rawPol.ExplainCode, rawPol.DefinedIn);
            pol.DisplayName        = ResolveString(rawPol.DisplayCode, rawPol.DefinedIn);

            if (!string.IsNullOrEmpty(rawPol.PresentationID))
                pol.Presentation = ResolvePresentation(rawPol.PresentationID, rawPol.DefinedIn);

            pol.UniqueID  = QualifyName(rawPol.ID, rawPol.DefinedIn);
            pol.RawPolicy = rawPol;
            polIds.Add(pol.UniqueID, pol);
        }

        // Second pass: Resolve references and link structures
        foreach (var cat in catIds.Values)
        {
            if (!string.IsNullOrEmpty(cat.RawCategory.ParentID))
            {
                var parentCatName = ResolveRef(cat.RawCategory.ParentID, cat.RawCategory.DefinedIn);
                var parentCat     = FindCatById(parentCatName);

                if (parentCat is null)
                    continue; // In case the parent category doesn't exist

                parentCat.Children.Add(cat);
                cat.Parent = parentCat;
            }
        }

        foreach (var product in productIds.Values)
        {
            if (product.RawProduct.Parent is not null)
            {
                var parentProductId = QualifyName(
                    product.RawProduct.Parent.ID,
                    product.RawProduct.DefinedIn
                ); // Child products can't be defined in other files

                var parentProduct = FindProductById(parentProductId);
                parentProduct.Children.Add(product);
                product.Parent = parentProduct;
            }
        }

        foreach (var sup in supIds.Values)
        {
            foreach (var supEntry in sup.Elements)
            {
                var targetId = ResolveRef(supEntry.RawSupportEntry.ProductID, sup.RawSupport.DefinedIn); // Support or product
                supEntry.Product = FindProductById(targetId);

                if (supEntry.Product is null)
                    supEntry.SupportDefinition = FindSupById(targetId);
            }
        }

        foreach (var pol in polIds.Values)
        {
            var catId    = ResolveRef(pol.RawPolicy.CategoryID, pol.RawPolicy.DefinedIn);
            var ownerCat = FindCatById(catId);

            if (ownerCat is not null)
            {
                ownerCat.Policies.Add(pol);
                pol.Category = ownerCat;
            }

            var supportId = ResolveRef(pol.RawPolicy.SupportedCode, pol.RawPolicy.DefinedIn);
            pol.SupportedOn = FindSupById(supportId);
        }

        // Third pass: Add items to the final lists
        foreach (var cat in catIds)
        {
            FlatCategories.Add(cat.Key, cat.Value);

            if (cat.Value.Parent is null)
                Categories.Add(cat.Key, cat.Value);
        }

        foreach (var product in productIds)
        {
            FlatProducts.Add(product.Key, product.Value);

            if (product.Value.Parent is null)
                Products.Add(product.Key, product.Value);
        }

        foreach (var pol in polIds)
            Policies.Add(pol.Key, pol.Value);

        foreach (var sup in supIds)
            SupportDefinitions.Add(sup.Key, sup.Value);

        // Purge the temporary partially-constructed items
        _rawCategories.Clear();
        _rawProducts.Clear();
        _rawSupport.Clear();
        _rawPolicies.Clear();
    }

    private T FindInTempOrFlat<T>(string uniqueId, Dictionary<string, T> tempDict, Dictionary<string, T> flatDict)
    {
        // Get the best available structure for an ID
        if (tempDict.ContainsKey(uniqueId))
            return tempDict[uniqueId];

        if (flatDict is not null && flatDict.ContainsKey(uniqueId))
            return flatDict[uniqueId];

        return default;
    }

    public string ResolveString(string displayCode, AdmxFile admx)
    {
        // Find a localized string from a display code
        if (string.IsNullOrEmpty(displayCode))
            return "";

        if (!displayCode.StartsWith("$(string."))
            return displayCode;

        var stringId = displayCode.Substring(9, displayCode.Length - 10);
        var dict     = _sourceFiles[admx].StringTable;

        if (dict.ContainsKey(stringId))
            return dict[stringId];

        return displayCode;
    }

    public Presentation ResolvePresentation(string displayCode, AdmxFile admx)
    {
        // Find a presentation from a code
        if (!displayCode.StartsWith("$(presentation."))
            return null;

        var presId = displayCode.Substring(15, displayCode.Length - 16);
        var dict   = _sourceFiles[admx].PresentationTable;

        if (dict.ContainsKey(presId))
            return dict[presId];

        return null;
    }

    private string QualifyName(string id, AdmxFile admx) => admx.AdmxNamespace + ":" + id;

    private string ResolveRef(string @ref, AdmxFile admx)
    {
        // Get a fully qualified name from a code and the current scope
        if (@ref.Contains(":"))
        {
            var parts = @ref.Split(new[] { ':' }, 2);

            if (admx.Prefixes.ContainsKey(parts[0]))
            {
                var srcNamespace = admx.Prefixes[parts[0]];

                return srcNamespace + ":" + parts[1];
            }

            return @ref;
            // Assume a literal
        }

        return QualifyName(@ref, admx);
    }
}

public enum AdmxLoadFailType
{
    BadAdmxParse,
    BadAdmx,
    NoAdml,
    BadAdmlParse,
    BadAdml,
    DuplicateNamespace
}

public class AdmxLoadFailure
{
    public string           AdmxPath;
    public AdmxLoadFailType FailType;
    public string           Info;

    public AdmxLoadFailure(AdmxLoadFailType FailType, string AdmxPath, string Info)
    {
        this.FailType = FailType;
        this.AdmxPath = AdmxPath;
        this.Info     = Info;
    }

    public AdmxLoadFailure(AdmxLoadFailType FailType, string AdmxPath) : this(FailType, AdmxPath, "") { }

    public override string ToString()
    {
        var failMsg = "Couldn't load " + AdmxPath + ": " + GetFailMessage(FailType, Info);

        if (!failMsg.EndsWith("."))
            failMsg += ".";

        return failMsg;
    }

    private static string GetFailMessage(AdmxLoadFailType FailType, string Info)
    {
        switch (FailType)
        {
            case AdmxLoadFailType.BadAdmxParse: { return "The ADMX XML couldn't be parsed: " + Info; }

            case AdmxLoadFailType.BadAdmx: { return "The ADMX is invalid: " + Info; }

            case AdmxLoadFailType.NoAdml: { return "The corresponding ADML is missing"; }

            case AdmxLoadFailType.BadAdmlParse: { return "The ADML XML couldn't be parsed: " + Info; }

            case AdmxLoadFailType.BadAdml: { return "The ADML is invalid: " + Info; }

            case AdmxLoadFailType.DuplicateNamespace: { return "The " + Info + " namespace is already owned by a different ADMX file"; }
        }

        return string.IsNullOrEmpty(Info) ? "An unknown error occurred" : Info;
    }
}