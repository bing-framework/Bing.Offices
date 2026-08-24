using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 映射配置加载器。
/// </summary>
public static class ExcelMappingConfigurationLoader
{
    private const int MaxDocumentBytes = 1024 * 1024;
    private const int MaxDepth = 32;
    private const int MaxColumns = 1000;
    private const int MaxAliasesPerColumn = 100;
    private const int MaxValidationsPerColumn = 100;
    private const int MaxStringLength = 4096;

    /// <summary>
    /// 从 JSON 文本加载兼容映射配置；v2 文档返回 Import 方向配置，v1 平铺文档原样迁移。
    /// </summary>
    public static ExcelMappingConfiguration FromJson(string json) => FromJsonDocument(json).Import;

    /// <summary>
    /// 从 JSON 流加载兼容映射配置。
    /// </summary>
    public static ExcelMappingConfiguration FromJson(Stream source) => FromJsonDocument(source).Import;

    /// <summary>
    /// 从 UTF-8 JSON 配置文件加载兼容映射配置。
    /// </summary>
    public static ExcelMappingConfiguration FromJsonFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("JSON 配置文件路径不能为空。", nameof(path));
        using var source = File.OpenRead(path);
        return FromJsonDocument(source).Import;
    }

    /// <summary>
    /// 从 JSON 文本加载规范化映射文档。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(string json)
        => LoadJsonDocument(json, null, null);

    /// <summary>
    /// 从 JSON 文本加载文档，并返回非阻断的迁移诊断。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(string json,
        out IReadOnlyList<ExcelMappingDiagnostic> diagnostics)
    {
        var items = new List<ExcelMappingDiagnostic>();
        var result = LoadJsonDocument(json, null, items);
        diagnostics = items;
        return result;
    }

    /// <summary>
    /// 从 JSON 文本加载文档，并按已注册的业务模型别名进行校验。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(string json, ExcelModelAliasRegistry modelAliases)
        => LoadJsonDocument(json, modelAliases, null);

    private static ExcelMappingDocument LoadJsonDocument(string json, ExcelModelAliasRegistry modelAliases,
        ICollection<ExcelMappingDiagnostic> diagnostics)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("JSON 配置不能为空。", nameof(json));
        if (Encoding.UTF8.GetByteCount(json) > MaxDocumentBytes)
            throw new InvalidOperationException($"JSON 配置超过最大字节数: {MaxDocumentBytes}");
        try
        {
            using var document = JsonDocument.Parse(json, new JsonDocumentOptions
            {
                MaxDepth = MaxDepth,
                CommentHandling = JsonCommentHandling.Disallow,
                AllowTrailingCommas = false
            });
            if (document.RootElement.ValueKind != JsonValueKind.Object)
                throw new InvalidOperationException("JSON 配置根节点必须是对象。");
            var isV2 = document.RootElement.TryGetProperty("version", out _)
                       || document.RootElement.TryGetProperty("import", out _)
                       || document.RootElement.TryGetProperty("export", out _);
            ValidateJsonElement(document.RootElement, "$", isV2);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                MaxDepth = MaxDepth
            };
            ExcelMappingDocument result;
            if (isV2)
            {
                result = JsonSerializer.Deserialize<ExcelMappingDocument>(json, options)
                    ?? throw new InvalidOperationException("JSON 配置未包含有效映射文档。");
                if (result.Version != 2)
                    throw new InvalidOperationException($"不支持的 JSON 映射文档版本: {result.Version}");
            }
            else
            {
                var configuration = JsonSerializer.Deserialize<ExcelMappingConfiguration>(json, options)
                    ?? throw new InvalidOperationException("JSON 配置未包含有效映射。");
                result = CreateV1Document(configuration);
                diagnostics?.Add(new ExcelMappingDiagnostic("V1_MIGRATED", "$",
                    "检测到 v1 平铺 JSON，已归一化为 v2 文档。"));
            }
            ValidateDocument(result, modelAliases);
            return result;
        }
        catch (JsonException exception)
        {
            throw new InvalidOperationException($"JSON 映射配置无效: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// 从 JSON 流加载规范化映射文档，且不关闭调用方流。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(Stream source)
        => FromJsonDocument(source, null);

    /// <summary>
    /// 从 JSON 流加载文档，并按已注册的业务模型别名进行校验。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(Stream source, ExcelModelAliasRegistry modelAliases)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("JSON 配置流不可读取。", nameof(source));
        using var reader = new StreamReader(source, Encoding.UTF8, true, 1024, true);
        return FromJsonDocument(ReadLimitedText(reader), modelAliases);
    }

    /// <summary>
    /// 从 XML 文本加载兼容映射配置；v2 文档返回 Import 方向配置。
    /// </summary>
    public static ExcelMappingConfiguration FromXml(string xml) => FromXmlDocument(xml).Import;

    /// <summary>
    /// 从 XML 流加载兼容映射配置。
    /// </summary>
    public static ExcelMappingConfiguration FromXml(Stream source) => FromXmlDocument(source).Import;

    /// <summary>
    /// 从 UTF-8 XML 配置文件加载兼容映射配置。
    /// </summary>
    public static ExcelMappingConfiguration FromXmlFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("XML 配置文件路径不能为空。", nameof(path));
        using var source = File.OpenRead(path);
        return FromXmlDocument(source).Import;
    }

    /// <summary>
    /// 从 XML 文本加载规范化映射文档。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(string xml)
        => LoadXmlDocument(xml, null, null);

    /// <summary>
    /// 从 XML 文本加载文档，并返回非阻断的迁移诊断。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(string xml,
        out IReadOnlyList<ExcelMappingDiagnostic> diagnostics)
    {
        var items = new List<ExcelMappingDiagnostic>();
        var result = LoadXmlDocument(xml, null, items);
        diagnostics = items;
        return result;
    }

    /// <summary>
    /// 从 XML 文本加载文档，并按已注册的业务模型别名进行校验。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(string xml, ExcelModelAliasRegistry modelAliases)
        => LoadXmlDocument(xml, modelAliases, null);

    private static ExcelMappingDocument LoadXmlDocument(string xml, ExcelModelAliasRegistry modelAliases,
        ICollection<ExcelMappingDiagnostic> diagnostics)
    {
        if (string.IsNullOrWhiteSpace(xml))
            throw new ArgumentException("XML 配置不能为空。", nameof(xml));
        if (Encoding.UTF8.GetByteCount(xml) > MaxDocumentBytes)
            throw new InvalidOperationException($"XML 配置超过最大字节数: {MaxDocumentBytes}");
        var isV2 = IsXmlDocumentRoot(xml);
        ValidateXmlShape(xml, isV2);
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        var result = DeserializeXml(reader);
        if (!isV2)
            diagnostics?.Add(new ExcelMappingDiagnostic("V1_MIGRATED", "/ExcelMappingConfiguration",
                "检测到 v1 平铺 XML，已归一化为 v2 文档。"));
        ValidateDocument(result, modelAliases);
        return result;
    }

    /// <summary>
    /// 从 XML 流加载规范化映射文档，且不关闭调用方流。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(Stream source)
        => FromXmlDocument(source, null);

    /// <summary>
    /// 从 XML 流加载文档，并按已注册的业务模型别名进行校验。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(Stream source, ExcelModelAliasRegistry modelAliases)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("XML 配置流不可读取。", nameof(source));
        using var reader = new StreamReader(source, Encoding.UTF8, true, 1024, true);
        return FromXmlDocument(ReadLimitedText(reader), modelAliases);
    }

    /// <summary>
    /// 将 normalized v2 文档写为 JSON。
    /// </summary>
    public static string ToJson(ExcelMappingDocument document)
    {
        ValidateDocument(document, null);
        return JsonSerializer.Serialize(document, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            IgnoreNullValues = true,
            WriteIndented = true
        });
    }

    /// <summary>
    /// 将 normalized v2 文档写为 XML。
    /// </summary>
    public static string ToXml(ExcelMappingDocument document)
    {
        ValidateDocument(document, null);
        var serializer = new XmlSerializer(typeof(ExcelMappingDocument));
        using var writer = new Utf8StringWriter();
        serializer.Serialize(writer, document);
        return writer.ToString();
    }

    private static string ReadLimitedText(TextReader reader)
    {
        var buffer = new char[8192];
        var builder = new StringBuilder();
        var total = 0;
        int count;
        while ((count = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            total += count;
            if (total > MaxDocumentBytes)
                throw new InvalidOperationException($"配置超过最大字符数: {MaxDocumentBytes}");
            builder.Append(buffer, 0, count);
        }
        return builder.ToString();
    }

    private static ExcelMappingDocument CreateV1Document(ExcelMappingConfiguration configuration) => new()
    {
        Version = 2,
        Import = configuration ?? throw new InvalidOperationException("JSON 配置未包含有效映射。"),
        Export = new ExcelMappingConfiguration()
    };

    private static XmlReaderSettings CreateXmlReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = MaxDocumentBytes,
        MaxCharactersFromEntities = 0
    };

    private static bool IsXmlDocumentRoot(string xml)
    {
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        reader.MoveToContent();
        return string.Equals(reader.LocalName, nameof(ExcelMappingDocument), StringComparison.Ordinal);
    }

    private static void ValidateXmlShape(string xml, bool isV2)
    {
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        var document = XDocument.Load(reader, LoadOptions.SetLineInfo);
        if (document.Root == null)
            throw new InvalidOperationException("XML 配置根节点不能为空。");
        var expectedRoot = isV2 ? nameof(ExcelMappingDocument) : nameof(ExcelMappingConfiguration);
        if (!string.Equals(document.Root.Name.LocalName, expectedRoot, StringComparison.Ordinal))
            throw new InvalidOperationException($"XML 根节点必须是 /{expectedRoot}。");
        ValidateXmlElement(document.Root, "/" + expectedRoot);
    }

    private static void ValidateXmlElement(XElement element, string path)
    {
        foreach (var attribute in element.Attributes())
        {
            if (attribute.IsNamespaceDeclaration || attribute.Name.NamespaceName == "http://www.w3.org/2001/XMLSchema-instance")
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

    private static bool IsKnownXmlElement(string parent, string child)
    {
        if (parent == nameof(ExcelMappingDocument))
            return child is nameof(ExcelMappingDocument.Version) or nameof(ExcelMappingDocument.Profile)
                or nameof(ExcelMappingDocument.ModelAlias) or nameof(ExcelMappingDocument.TenantId)
                or nameof(ExcelMappingDocument.ConfigurationVersion) or nameof(ExcelMappingDocument.Import)
                or nameof(ExcelMappingDocument.Export);
        if (parent is nameof(ExcelMappingDocument.Import) or nameof(ExcelMappingDocument.Export)
            or nameof(ExcelMappingConfiguration))
            return child is nameof(ExcelMappingConfiguration.SourceKind) or nameof(ExcelMappingConfiguration.Columns)
                or nameof(ExcelMappingConfiguration.DynamicColumns) or nameof(ExcelMappingConfiguration.Style)
                or nameof(ExcelMappingConfiguration.Layout);
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
                or nameof(ExcelMappingStyleConfiguration.BodyStyleKey)
                or nameof(ExcelMappingStyleConfiguration.NumberFormat);
        if (parent == nameof(ExcelMappingConfiguration.Layout))
            return child is nameof(ExcelMappingLayoutConfiguration.ColumnIndex)
                or nameof(ExcelMappingLayoutConfiguration.PlacementKey);
        if (parent == nameof(ExcelColumnConfiguration))
            return child is nameof(ExcelColumnConfiguration.PropertyName) or nameof(ExcelColumnConfiguration.Title)
                or nameof(ExcelColumnConfiguration.Aliases) or nameof(ExcelColumnConfiguration.ColumnIndex)
                or nameof(ExcelColumnConfiguration.Ignored) or nameof(ExcelColumnConfiguration.Formatter)
                or nameof(ExcelColumnConfiguration.DecimalScale) or nameof(ExcelColumnConfiguration.ConverterName)
                or nameof(ExcelColumnConfiguration.ImportWhitespace)
                or nameof(ExcelColumnConfiguration.ValidationRuleNames)
                or nameof(ExcelColumnConfiguration.ValidationRuleNamesToRemove)
                or nameof(ExcelColumnConfiguration.ClearValidationRules)
                or nameof(ExcelColumnConfiguration.ValidationRuleMergeMode)
                or nameof(ExcelColumnConfiguration.ValueMappings)
                or nameof(ExcelColumnConfiguration.ValueMappingMergeMode)
                or nameof(ExcelColumnConfiguration.ImageMultiplicity);
        if (parent is nameof(ExcelColumnConfiguration.Aliases)
            or nameof(ExcelColumnConfiguration.ValidationRuleNames)
            or nameof(ExcelColumnConfiguration.ValidationRuleNamesToRemove))
            return child == "string";
        if (parent == nameof(ExcelMappingDynamicColumnConfiguration.Aliases))
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

    private static ExcelMappingDocument DeserializeXml(XmlReader reader)
    {
        try
        {
            reader.MoveToContent();
            if (string.Equals(reader.LocalName, nameof(ExcelMappingDocument), StringComparison.Ordinal))
            {
                var serializer = new XmlSerializer(typeof(ExcelMappingDocument));
                AttachXmlValidationHandlers(serializer);
                return (ExcelMappingDocument)serializer.Deserialize(reader)
                    ?? throw new InvalidOperationException("XML 配置未包含有效映射文档。");
            }
            var configurationSerializer = new XmlSerializer(typeof(ExcelMappingConfiguration));
            AttachXmlValidationHandlers(configurationSerializer);
            var configuration = (ExcelMappingConfiguration)configurationSerializer.Deserialize(reader);
            return CreateV1Document(configuration);
        }
        catch (InvalidOperationException exception)
        {
            var validationException = FindInnerException<XmlMappingValidationException>(exception);
            if (validationException != null)
                throw new InvalidOperationException(validationException.Message, validationException);
            if (exception.InnerException is XmlException xmlException)
                throw xmlException;
            throw;
        }
    }

    private static TException FindInnerException<TException>(Exception exception)
        where TException : Exception
    {
        while (exception != null)
        {
            if (exception is TException match)
                return match;
            exception = exception.InnerException;
        }
        return null;
    }

    private static void ValidateDocument(ExcelMappingDocument document, ExcelModelAliasRegistry modelAliases)
    {
        if (document == null)
            throw new InvalidOperationException("映射文档不能为空。");
        if (document.Version != 2)
            throw new InvalidOperationException($"不支持的映射文档版本: {document.Version}");
        ValidateBusinessAlias(document.Profile, "profile");
        ValidateBusinessAlias(document.ModelAlias, "modelAlias");
        ValidateText(document.TenantId, "tenantId");
        ValidateText(document.ConfigurationVersion, "configurationVersion");
        if (modelAliases != null && modelAliases.HasRegistrations
            && !string.IsNullOrWhiteSpace(document.ModelAlias)
            && !modelAliases.Contains(document.ModelAlias))
            throw new InvalidOperationException($"未知 modelAlias: {document.ModelAlias}");
        ValidateConfiguration(document.Import, "import");
        ValidateConfiguration(document.Export, "export");
        ValidateText(document.ModelAlias, "modelAlias");
    }

    private static void ValidateConfiguration(ExcelMappingConfiguration configuration, string path)
    {
        if (configuration == null)
            return;
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
        for (var dynamicIndex = 0; dynamicIndex < (configuration.DynamicColumns?.Count ?? 0); dynamicIndex++)
        {
            var dynamic = configuration.DynamicColumns[dynamicIndex];
            var dynamicPath = $"{path}.dynamicColumns[{dynamicIndex}]";
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
        for (var columnIndex = 0; columnIndex < configuration.Columns.Count; columnIndex++)
        {
            var column = configuration.Columns[columnIndex];
            var columnPath = $"{path}.columns[{columnIndex}]";
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
            for (var aliasIndex = 0; aliasIndex < (column.Aliases?.Count ?? 0); aliasIndex++)
                ValidateText(column.Aliases[aliasIndex], $"{columnPath}.aliases[{aliasIndex}]");
            for (var ruleIndex = 0; ruleIndex < (column.ValidationRuleNames?.Count ?? 0); ruleIndex++)
                ValidateText(column.ValidationRuleNames[ruleIndex], $"{columnPath}.validationRuleNames[{ruleIndex}]");
            for (var ruleIndex = 0; ruleIndex < (column.ValidationRuleNamesToRemove?.Count ?? 0); ruleIndex++)
                ValidateText(column.ValidationRuleNamesToRemove[ruleIndex], $"{columnPath}.validationRuleNamesToRemove[{ruleIndex}]");
            for (var mappingIndex = 0; mappingIndex < (column.ValueMappings?.Count ?? 0); mappingIndex++)
            {
                var mapping = column.ValueMappings[mappingIndex];
                ValidateText(mapping?.Text, $"{columnPath}.valueMappings[{mappingIndex}].text");
                ValidateText(mapping?.Value, $"{columnPath}.valueMappings[{mappingIndex}].value");
            }
        }
    }

    private static void ValidateJsonElement(JsonElement element, string path, bool isV2)
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

    private static bool IsKnownJsonProperty(string path, string propertyName, bool isV2)
    {
        var names = path == "$"
            ? isV2
                ? new[] { "version", "profile", "modelAlias", "tenantId", "configurationVersion", "import", "export" }
                : new[] { "sourceKind", "columns" }
            : path.EndsWith(".import", StringComparison.Ordinal) || path.EndsWith(".export", StringComparison.Ordinal)
                ? new[] { "sourceKind", "columns", "dynamicColumns", "style", "layout" }
                : path.EndsWith(".dynamicColumns", StringComparison.Ordinal)
                    ? new[] { "key", "title", "aliases", "dataTypeName", "order", "converterName", "validatorName", "validationRuleNames", "validationRules", "numberFormat", "columnIndex", "placementKey", "imageMultiplicity" }
                    : path.EndsWith(".validationRules", StringComparison.Ordinal)
                        ? new[] { "name", "pattern", "format", "cultureName", "min", "max", "maxValue", "maxLength", "ignoreEmpty" }
                    : path.Contains(".validationRules[", StringComparison.Ordinal)
                        ? new[] { "name", "pattern", "format", "cultureName", "min", "max", "maxValue", "maxLength", "ignoreEmpty" }
                    : path.Contains(".dynamicColumns[", StringComparison.Ordinal)
                        ? new[] { "key", "title", "aliases", "dataTypeName", "order", "converterName", "validatorName", "validationRuleNames", "validationRules", "numberFormat", "columnIndex", "placementKey", "imageMultiplicity" }
                        : path.EndsWith(".style", StringComparison.Ordinal)
                            ? new[] { "headerStyleKey", "bodyStyleKey", "numberFormat" }
                            : path.EndsWith(".layout", StringComparison.Ordinal)
                                ? new[] { "columnIndex", "placementKey" }
                                    : path.Contains(".dynamicColumns[", StringComparison.Ordinal)
                                        && path.EndsWith(".aliases", StringComparison.Ordinal)
                                        ? new[] { "value" }
                : path.EndsWith(".valueMappings", StringComparison.Ordinal) || path.Contains(".valueMappings[", StringComparison.Ordinal)
                    ? new[] { "text", "value" }
                    : path.EndsWith(".columns", StringComparison.Ordinal) || path.Contains(".columns[", StringComparison.Ordinal)
                        ? new[] { "propertyName", "title", "aliases", "columnIndex", "ignored", "formatter", "decimalScale", "converterName", "importWhitespace", "validationRuleNames", "validationRuleNamesToRemove", "clearValidationRules", "validationRuleMergeMode", "valueMappings", "valueMappingMergeMode", "imageMultiplicity" }
                        : Array.Empty<string>();
        return names.Any(name => string.Equals(name, propertyName, StringComparison.OrdinalIgnoreCase));
    }

    private static void AttachXmlValidationHandlers(XmlSerializer serializer)
    {
        serializer.UnknownNode += (_, eventArgs) =>
        throw new XmlMappingValidationException($"未知 XML 字段: /ExcelMappingDocument/{eventArgs.Name}");
        serializer.UnknownAttribute += (_, eventArgs) =>
        throw new XmlMappingValidationException($"未知 XML 属性: /ExcelMappingDocument/@{eventArgs.Attr?.Name ?? eventArgs.Attr?.LocalName}");
    }

    private sealed class XmlMappingValidationException : InvalidOperationException
    {
        public XmlMappingValidationException(string message) : base(message)
        {
        }
    }

    private static void ValidateText(string value, string path)
    {
        if (value != null && value.Length > MaxStringLength)
            throw new InvalidOperationException($"{path} 字符串长度超过最大值: {MaxStringLength}");
    }

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

    private static void ValidateBusinessAlias(string value, string path)
    {
        if (string.IsNullOrWhiteSpace(value))
            return;
        if (value.Length > 256 || value.IndexOfAny(new[] { '.', ',', '+', '[', ']' }) >= 0
            || value.IndexOf("::", StringComparison.Ordinal) >= 0)
            throw new InvalidOperationException($"{path} 必须是稳定业务别名，不能使用 CLR 或程序集限定类型名。");
    }

    private sealed class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }
}

/// <summary>
/// 映射配置加载器的默认服务实现。
/// </summary>
public sealed class DefaultExcelMappingConfigurationLoader : IExcelMappingConfigurationLoader
{
    /// <inheritdoc />
    public ExcelMappingConfiguration FromJson(string json) => ExcelMappingConfigurationLoader.FromJson(json);

    /// <inheritdoc />
    public ExcelMappingConfiguration FromJson(Stream source) => ExcelMappingConfigurationLoader.FromJson(source);

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(string json) => ExcelMappingConfigurationLoader.FromJsonDocument(json);

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(Stream source) => ExcelMappingConfigurationLoader.FromJsonDocument(source);

    /// <inheritdoc />
    public ExcelMappingConfiguration FromXml(string xml) => ExcelMappingConfigurationLoader.FromXml(xml);

    /// <inheritdoc />
    public ExcelMappingConfiguration FromXml(Stream source) => ExcelMappingConfigurationLoader.FromXml(source);

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(string xml) => ExcelMappingConfigurationLoader.FromXmlDocument(xml);

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(Stream source) => ExcelMappingConfigurationLoader.FromXmlDocument(source);
}
