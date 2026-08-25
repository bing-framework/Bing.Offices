using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;

namespace Bing.Offices.Configurations;

/// <summary>
/// 负责验证 JSON/XML 映射文档的字段结构和业务约束。
/// </summary>
internal static class ExcelMappingDocumentValidator
{
    private const int MaxColumns = 1000;
    private const int MaxAliasesPerColumn = 100;
    private const int MaxValidationsPerColumn = 100;
    private const int MaxStringLength = 4096;

    /// <summary>
    /// 验证 JSON 元素名称和值的结构边界。
    /// </summary>
    internal static void ValidateJsonElement(JsonElement element, string path, bool isV2)
    {
        if (element.ValueKind == JsonValueKind.String)
        {
            ValidateText(element.GetString(), path);
            return;
        }
        if (element.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in element.EnumerateArray())
                ValidateJsonElement(item, $"{path}[{index++}]", isV2);
            return;
        }
        if (element.ValueKind != JsonValueKind.Object)
            return;
        foreach (var property in element.EnumerateObject())
        {
            var propertyPath = path == "$" ? $"$.{property.Name}" : $"{path}.{property.Name}";
            if (!IsKnownJsonProperty(path, property.Name, isV2))
                throw new InvalidOperationException($"未知 JSON 字段: {propertyPath}");
            ValidateJsonElement(property.Value, propertyPath, isV2);
        }
    }

    /// <summary>
    /// 验证 XML 节点名称、属性和层级结构。
    /// </summary>
    internal static void ValidateXmlShape(XElement root, bool isV2)
    {
        if (root == null)
            throw new InvalidOperationException("XML 配置根节点不能为空。");
        var expectedRoot = isV2 ? nameof(ExcelMappingDocument) : nameof(ExcelMappingConfiguration);
        if (!string.Equals(root.Name.LocalName, expectedRoot, StringComparison.Ordinal))
            throw new InvalidOperationException($"XML 根节点必须是 /{expectedRoot}。");
        ValidateXmlElement(root, "/" + expectedRoot);
    }

    /// <summary>
    /// 验证映射文档的版本、方向配置和业务字段限制。
    /// </summary>
    internal static void ValidateDocument(ExcelMappingDocument document, ExcelModelAliasRegistry modelAliases)
    {
        if (document == null)
            throw new InvalidOperationException("映射文档不能为空。");
        if (document.Version != 2)
            throw new InvalidOperationException($"不支持的映射文档版本: {document.Version}");
        ValidateText(document.TenantId, "tenantId");
        ValidateText(document.ConfigurationVersion, "configurationVersion");
        ValidateConfiguration(document.Import, "import", modelAliases);
        ValidateConfiguration(document.Export, "export", modelAliases);
    }

    /// <summary>
    /// 验证单个方向配置及其列、动态列和校验规则限制。
    /// </summary>
    private static void ValidateConfiguration(ExcelMappingConfiguration configuration, string path,
        ExcelModelAliasRegistry modelAliases)
    {
        if (configuration == null)
            return;
        if (configuration.DynamicColumnMergeMode.HasValue
            && !Enum.IsDefined(typeof(ExcelDynamicColumnMergeMode), configuration.DynamicColumnMergeMode.Value))
            throw new InvalidOperationException($"{path}.dynamicColumnMergeMode 无效。");
        for (var index = 0; index < (configuration.DynamicColumnKeysToRemove?.Count ?? 0); index++)
            ValidateText(configuration.DynamicColumnKeysToRemove[index],
                $"{path}.dynamicColumnKeysToRemove[{index}]");
        ValidateBusinessAlias(configuration.Profile, $"{path}.profile");
        ValidateBusinessAlias(configuration.ModelAlias, $"{path}.modelAlias");
        if (modelAliases != null && modelAliases.HasRegistrations
            && !string.IsNullOrWhiteSpace(configuration.ModelAlias)
            && !modelAliases.Contains(configuration.ModelAlias))
            throw new InvalidOperationException($"未知 modelAlias: {configuration.ModelAlias}");
        ValidateText(configuration.Style?.HeaderStyleKey, $"{path}.style.headerStyleKey");
        ValidateText(configuration.Style?.BodyStyleKey, $"{path}.style.bodyStyleKey");
        ValidateText(configuration.Style?.NumberFormat, $"{path}.style.numberFormat");
        ValidatePlacementKey(configuration.Layout?.PlacementKey, $"{path}.layout.placementKey");
        if (configuration.Layout?.ColumnIndex < 0)
            throw new InvalidOperationException($"{path}.layout.columnIndex 不能小于 0。");
        if (configuration.Columns == null)
            throw new InvalidOperationException($"{path}.columns 不能为空。");
        if (configuration.Columns.Count > MaxColumns)
            throw new InvalidOperationException($"{path}.columns 超过最大数量: {MaxColumns}");
        if ((configuration.DynamicColumns?.Count ?? 0) > MaxColumns)
            throw new InvalidOperationException($"{path}.dynamicColumns 超过最大数量: {MaxColumns}");

        for (var index = 0; index < (configuration.DynamicColumns?.Count ?? 0); index++)
        {
            var dynamic = configuration.DynamicColumns[index];
            var dynamicPath = $"{path}.dynamicColumns[{index}]";
            if (dynamic == null)
                throw new InvalidOperationException($"{dynamicPath} 不能为 null。");
            ValidateText(dynamic.Key, $"{dynamicPath}.key");
            ValidateText(dynamic.Title, $"{dynamicPath}.title");
            ValidateText(dynamic.DataTypeName, $"{dynamicPath}.dataTypeName");
            ValidateText(dynamic.ConverterName, $"{dynamicPath}.converterName");
            ValidateText(dynamic.ValidatorName, $"{dynamicPath}.validatorName");
            for (var validationIndex = 0; validationIndex < (dynamic.ValidationRules?.Count ?? 0); validationIndex++)
            {
                var validation = dynamic.ValidationRules[validationIndex];
                var validationPath = $"{dynamicPath}.validationRules[{validationIndex}]";
                if (validation == null)
                    throw new InvalidOperationException($"{validationPath} 不能为 null。");
                ValidateText(validation.Name, $"{validationPath}.name");
                ValidateText(validation.Pattern, $"{validationPath}.pattern");
                ValidateText(validation.Format, $"{validationPath}.format");
                ValidateText(validation.CultureName, $"{validationPath}.cultureName");
                if (validation.MaxLength < 0)
                    throw new InvalidOperationException($"{validationPath}.maxLength 不能小于 0。");
            }
            for (var ruleIndex = 0; ruleIndex < (dynamic.ValidationRuleNames?.Count ?? 0); ruleIndex++)
                ValidateText(dynamic.ValidationRuleNames[ruleIndex],
                    $"{dynamicPath}.validationRuleNames[{ruleIndex}]");
            ValidateText(dynamic.NumberFormat, $"{dynamicPath}.numberFormat");
            ValidatePlacementKey(dynamic.PlacementKey, $"{dynamicPath}.placementKey");
            if (dynamic.ColumnIndex < 0)
                throw new InvalidOperationException($"{dynamicPath}.columnIndex 不能小于 0。");
            if (dynamic.ColumnIndex.HasValue && !string.IsNullOrWhiteSpace(dynamic.PlacementKey))
                throw new InvalidOperationException($"{dynamicPath} 不能同时设置 columnIndex 和 placementKey。");
            for (var aliasIndex = 0; aliasIndex < (dynamic.Aliases?.Count ?? 0); aliasIndex++)
                ValidateText(dynamic.Aliases[aliasIndex], $"{dynamicPath}.aliases[{aliasIndex}]");
        }

        for (var index = 0; index < configuration.Columns.Count; index++)
        {
            var column = configuration.Columns[index];
            var columnPath = $"{path}.columns[{index}]";
            if (column == null)
                throw new InvalidOperationException($"{columnPath} 不能为 null。");
            if ((column.Aliases?.Count ?? 0) > MaxAliasesPerColumn)
                throw new InvalidOperationException($"{columnPath}.aliases 超过最大数量: {MaxAliasesPerColumn}");
            if ((column.ValidationRuleNames?.Count ?? 0) + (column.ValidationRuleNamesToRemove?.Count ?? 0)
                > MaxValidationsPerColumn)
                throw new InvalidOperationException($"{columnPath}.validations 超过最大数量: {MaxValidationsPerColumn}");
            ValidateText(column.PropertyName, $"{columnPath}.propertyName");
            ValidateText(column.Title, $"{columnPath}.title");
            ValidateText(column.Formatter, $"{columnPath}.formatter");
            ValidateText(column.ConverterName, $"{columnPath}.converterName");
            if (column.ValidationRuleMergeMode.HasValue
                && !Enum.IsDefined(typeof(ExcelValidationRuleMergeMode), column.ValidationRuleMergeMode.Value))
                throw new InvalidOperationException($"{columnPath}.validationRuleMergeMode 无效。");
            if (column.ValueMappingMergeMode.HasValue
                && !Enum.IsDefined(typeof(ExcelValueMappingMergeMode), column.ValueMappingMergeMode.Value))
                throw new InvalidOperationException($"{columnPath}.valueMappingMergeMode 无效。");
            for (var aliasIndex = 0; aliasIndex < (column.Aliases?.Count ?? 0); aliasIndex++)
                ValidateText(column.Aliases[aliasIndex], $"{columnPath}.aliases[{aliasIndex}]");
            for (var ruleIndex = 0; ruleIndex < (column.ValidationRuleNames?.Count ?? 0); ruleIndex++)
                ValidateText(column.ValidationRuleNames[ruleIndex],
                    $"{columnPath}.validationRuleNames[{ruleIndex}]");
            for (var ruleIndex = 0; ruleIndex < (column.ValidationRuleNamesToRemove?.Count ?? 0); ruleIndex++)
                ValidateText(column.ValidationRuleNamesToRemove[ruleIndex],
                    $"{columnPath}.validationRuleNamesToRemove[{ruleIndex}]");
            for (var mappingIndex = 0; mappingIndex < (column.ValueMappings?.Count ?? 0); mappingIndex++)
            {
                var mapping = column.ValueMappings[mappingIndex];
                ValidateText(mapping?.Text, $"{columnPath}.valueMappings[{mappingIndex}].text");
                ValidateText(mapping?.Value, $"{columnPath}.valueMappings[{mappingIndex}].value");
            }
        }
    }

    /// <summary>
    /// 验证单个 XML 节点的属性和已知子节点。
    /// </summary>
    private static void ValidateXmlElement(XElement element, string path)
    {
        foreach (var attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration
                || attribute.Name.NamespaceName == "http://www.w3.org/2001/XMLSchema-instance")
                continue;
            throw new InvalidOperationException($"未知 XML 属性: {path}/@{attribute.Name.LocalName}");
        }
        foreach (var child in element.Elements())
        {
            if (!IsKnownXmlElement(element.Name.LocalName, child.Name.LocalName))
                throw new InvalidOperationException($"未知 XML 字段: {path}/{child.Name.LocalName}");
            ValidateXmlElement(child, $"{path}/{child.Name.LocalName}");
        }
    }

    /// <summary>
    /// 判断 XML 子节点是否属于当前 schema。
    /// </summary>
    private static bool IsKnownXmlElement(string parent, string child)
    {
        if (parent == nameof(ExcelMappingDocument))
            return child is nameof(ExcelMappingDocument.Version) or nameof(ExcelMappingDocument.TenantId)
                or nameof(ExcelMappingDocument.ConfigurationVersion) or nameof(ExcelMappingDocument.UseConventionFallback)
                or nameof(ExcelMappingDocument.Import) or nameof(ExcelMappingDocument.Export);
        if (parent is nameof(ExcelMappingDocument.Import) or nameof(ExcelMappingDocument.Export)
            or nameof(ExcelMappingConfiguration))
            return child is nameof(ExcelMappingConfiguration.SourceKind) or nameof(ExcelMappingConfiguration.Profile)
                or nameof(ExcelMappingConfiguration.ModelAlias) or nameof(ExcelMappingConfiguration.Columns)
                or nameof(ExcelMappingConfiguration.DynamicColumns)
                or nameof(ExcelMappingConfiguration.DynamicColumnKeysToRemove)
                or nameof(ExcelMappingConfiguration.DynamicColumnMergeMode)
                or nameof(ExcelMappingConfiguration.Style) or nameof(ExcelMappingConfiguration.Layout)
                or nameof(ExcelMappingConfiguration.ClearDynamicColumns)
                or nameof(ExcelMappingConfiguration.ResetStyle) or nameof(ExcelMappingConfiguration.ResetLayout);
        if (parent == nameof(ExcelMappingConfiguration.Columns))
            return child == nameof(ExcelColumnConfiguration);
        if (parent == nameof(ExcelMappingConfiguration.DynamicColumns))
            return child == nameof(ExcelMappingDynamicColumnConfiguration);
        if (parent == nameof(ExcelMappingDynamicColumnConfiguration))
            return child is nameof(ExcelMappingDynamicColumnConfiguration.Key)
                or nameof(ExcelMappingDynamicColumnConfiguration.Title)
                or nameof(ExcelMappingDynamicColumnConfiguration.Aliases)
                or nameof(ExcelMappingDynamicColumnConfiguration.DataTypeName)
                or nameof(ExcelMappingDynamicColumnConfiguration.Order)
                or nameof(ExcelMappingDynamicColumnConfiguration.ConverterName)
                or nameof(ExcelMappingDynamicColumnConfiguration.ValidatorName)
                or nameof(ExcelMappingDynamicColumnConfiguration.ValidationRuleNames)
                or nameof(ExcelMappingDynamicColumnConfiguration.ValidationRules)
                or nameof(ExcelMappingDynamicColumnConfiguration.NumberFormat)
                or nameof(ExcelMappingDynamicColumnConfiguration.ColumnIndex)
                or nameof(ExcelMappingDynamicColumnConfiguration.PlacementKey)
                or nameof(ExcelMappingDynamicColumnConfiguration.ImageMultiplicity);
        if (parent == nameof(ExcelMappingConfiguration.Style))
            return child is nameof(ExcelMappingStyleConfiguration.HeaderStyleKey)
                or nameof(ExcelMappingStyleConfiguration.ClearHeaderStyleKey)
                or nameof(ExcelMappingStyleConfiguration.BodyStyleKey)
                or nameof(ExcelMappingStyleConfiguration.ClearBodyStyleKey)
                or nameof(ExcelMappingStyleConfiguration.NumberFormat)
                or nameof(ExcelMappingStyleConfiguration.ClearNumberFormat);
        if (parent == nameof(ExcelMappingConfiguration.Layout))
            return child is nameof(ExcelMappingLayoutConfiguration.ColumnIndex)
                or nameof(ExcelMappingLayoutConfiguration.ResetColumnIndex)
                or nameof(ExcelMappingLayoutConfiguration.PlacementKey)
                or nameof(ExcelMappingLayoutConfiguration.ClearPlacementKey);
        if (parent == nameof(ExcelColumnConfiguration))
            return child is nameof(ExcelColumnConfiguration.PropertyName) or nameof(ExcelColumnConfiguration.Title)
                or nameof(ExcelColumnConfiguration.Aliases) or nameof(ExcelColumnConfiguration.ColumnIndex)
                or nameof(ExcelColumnConfiguration.Ignored) or nameof(ExcelColumnConfiguration.Formatter)
                or nameof(ExcelColumnConfiguration.ClearTitle) or nameof(ExcelColumnConfiguration.ClearAliases)
                or nameof(ExcelColumnConfiguration.ResetColumnIndex) or nameof(ExcelColumnConfiguration.ResetIgnored)
                or nameof(ExcelColumnConfiguration.ClearFormatter) or nameof(ExcelColumnConfiguration.ResetDecimalScale)
                or nameof(ExcelColumnConfiguration.ClearConverterName)
                or nameof(ExcelColumnConfiguration.ResetImportWhitespace)
                or nameof(ExcelColumnConfiguration.DecimalScale) or nameof(ExcelColumnConfiguration.ConverterName)
                or nameof(ExcelColumnConfiguration.ImportWhitespace)
                or nameof(ExcelColumnConfiguration.ValidationRuleNames)
                or nameof(ExcelColumnConfiguration.ValidationRuleNamesToRemove)
                or nameof(ExcelColumnConfiguration.ClearValidationRules)
                or nameof(ExcelColumnConfiguration.ValidationRuleMergeMode)
                or nameof(ExcelColumnConfiguration.ValueMappings)
                or nameof(ExcelColumnConfiguration.ClearValueMappings)
                or nameof(ExcelColumnConfiguration.ValueMappingMergeMode)
                or nameof(ExcelColumnConfiguration.ResetImageMultiplicity)
                or nameof(ExcelColumnConfiguration.ImageMultiplicity);
        if (parent is nameof(ExcelColumnConfiguration.Aliases)
            or nameof(ExcelColumnConfiguration.ValidationRuleNames)
            or nameof(ExcelColumnConfiguration.ValidationRuleNamesToRemove)
            or nameof(ExcelMappingDynamicColumnConfiguration.Aliases))
            return child == "string";
        if (parent == nameof(ExcelMappingDynamicColumnConfiguration.ValidationRules))
            return child == nameof(ExcelMappingDynamicValidationConfiguration);
        if (parent == nameof(ExcelMappingDynamicValidationConfiguration))
            return child is nameof(ExcelMappingDynamicValidationConfiguration.Name)
                or nameof(ExcelMappingDynamicValidationConfiguration.Pattern)
                or nameof(ExcelMappingDynamicValidationConfiguration.Format)
                or nameof(ExcelMappingDynamicValidationConfiguration.CultureName)
                or nameof(ExcelMappingDynamicValidationConfiguration.Min)
                or nameof(ExcelMappingDynamicValidationConfiguration.Max)
                or nameof(ExcelMappingDynamicValidationConfiguration.MaxValue)
                or nameof(ExcelMappingDynamicValidationConfiguration.MaxLength)
                or nameof(ExcelMappingDynamicValidationConfiguration.IgnoreEmpty);
        if (parent == nameof(ExcelColumnConfiguration.ValueMappings))
            return child == nameof(ExcelValueMappingConfiguration);
        if (parent == nameof(ExcelValueMappingConfiguration))
            return child is nameof(ExcelValueMappingConfiguration.Text) or nameof(ExcelValueMappingConfiguration.Value);
        return true;
    }

    /// <summary>
    /// 判断 JSON 属性是否属于当前 schema。
    /// </summary>
    private static bool IsKnownJsonProperty(string path, string propertyName, bool isV2)
    {
        var names = path == "$"
            ? isV2
                ? new[] { "version", "tenantId", "configurationVersion", "useConventionFallback", "import", "export" }
                : new[] { "sourceKind", "profile", "modelAlias", "columns", "dynamicColumns", "dynamicColumnKeysToRemove", "dynamicColumnMergeMode", "style", "layout", "clearDynamicColumns", "resetStyle", "resetLayout" }
            : path.EndsWith(".import", StringComparison.Ordinal) || path.EndsWith(".export", StringComparison.Ordinal)
                ? new[] { "sourceKind", "profile", "modelAlias", "columns", "dynamicColumns", "dynamicColumnKeysToRemove", "dynamicColumnMergeMode", "style", "layout", "clearDynamicColumns", "resetStyle", "resetLayout" }
                : path.EndsWith(".dynamicColumns", StringComparison.Ordinal)
                    ? new[] { "key", "title", "aliases", "dataTypeName", "order", "converterName", "validatorName", "validationRuleNames", "validationRules", "numberFormat", "columnIndex", "placementKey", "imageMultiplicity" }
                    : path.EndsWith(".validationRules", StringComparison.Ordinal) || path.Contains(".validationRules[", StringComparison.Ordinal)
                        ? new[] { "name", "pattern", "format", "cultureName", "min", "max", "maxValue", "maxLength", "ignoreEmpty" }
                        : path.Contains(".dynamicColumns[", StringComparison.Ordinal)
                            ? new[] { "key", "title", "aliases", "dataTypeName", "order", "converterName", "validatorName", "validationRuleNames", "validationRules", "numberFormat", "columnIndex", "placementKey", "imageMultiplicity" }
                            : path.EndsWith(".style", StringComparison.Ordinal)
                                ? new[] { "headerStyleKey", "clearHeaderStyleKey", "bodyStyleKey", "clearBodyStyleKey", "numberFormat", "clearNumberFormat" }
                                : path.EndsWith(".layout", StringComparison.Ordinal)
                                    ? new[] { "columnIndex", "resetColumnIndex", "placementKey", "clearPlacementKey" }
                                    : path.EndsWith(".valueMappings", StringComparison.Ordinal) || path.Contains(".valueMappings[", StringComparison.Ordinal)
                                        ? new[] { "text", "value" }
                                        : path.EndsWith(".columns", StringComparison.Ordinal) || path.Contains(".columns[", StringComparison.Ordinal)
                                            ? new[] { "propertyName", "title", "clearTitle", "aliases", "clearAliases", "columnIndex", "resetColumnIndex", "ignored", "resetIgnored", "formatter", "clearFormatter", "decimalScale", "resetDecimalScale", "converterName", "clearConverterName", "importWhitespace", "resetImportWhitespace", "validationRuleNames", "validationRuleNamesToRemove", "clearValidationRules", "validationRuleMergeMode", "valueMappings", "clearValueMappings", "valueMappingMergeMode", "imageMultiplicity", "resetImageMultiplicity" }
                                            : Array.Empty<string>();
        return names.Any(name => string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// 验证文本字段长度。
    /// </summary>
    private static void ValidateText(string value, string path)
    {
        if (value != null && value.Length > MaxStringLength)
            throw new InvalidOperationException($"{path} 字符串长度超过最大值: {MaxStringLength}");
    }

    /// <summary>
    /// 验证动态列相对位置键。
    /// </summary>
    private static void ValidatePlacementKey(string value, string path)
    {
        ValidateText(value, path);
        if (string.IsNullOrWhiteSpace(value))
            return;
        var valid = value.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || value.StartsWith("after:", StringComparison.OrdinalIgnoreCase)
            || value.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            || value.StartsWith("after-", StringComparison.OrdinalIgnoreCase);
        var separator = value.IndexOfAny(new[] { ':', '-' });
        if (!valid || separator == value.Length - 1)
            throw new InvalidOperationException($"{path} 必须使用 before:列键 或 after:列键。");
    }

    /// <summary>
    /// 验证业务别名不是 CLR 或程序集限定类型名。
    /// </summary>
    private static void ValidateBusinessAlias(string value, string path)
    {
        if (string.IsNullOrWhiteSpace(value))
            return;
        if (value.Length > 256 || value.IndexOfAny(new[] { '.', ',', '+', '[', ']' }) >= 0
            || value.IndexOf("::", StringComparison.Ordinal) >= 0)
            throw new InvalidOperationException($"{path} 必须是稳定业务别名，不能使用 CLR 或程序集限定类型名。");
    }
}
