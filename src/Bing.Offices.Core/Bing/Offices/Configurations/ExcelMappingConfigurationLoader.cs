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
    private const int MaxDocumentBytes = ExcelMappingTextReader.MaxDocumentBytes;
    private const int MaxDepth = 32;
    private const int MaxColumns = 1000;
    private const int MaxAliasesPerColumn = 100;
    private const int MaxValidationsPerColumn = 100;
    private const int MaxStringLength = 4096;

    /// <summary>
    /// 从 JSON 文本加载规范化映射文档。
    /// </summary>
    public static ExcelMappingDocument FromJsonDocument(string json)
        => LoadJsonDocument(json, null, null);


    /// <summary>
    /// 将 v1 平铺 JSON 映射配置迁移到指定方向的 v2 文档。
    /// </summary>
    /// <param name="json">v1 JSON 配置。</param>
    /// <param name="direction">迁移目标方向。</param>
    public static ExcelMappingDocument MigrateV1Json(string json, MappingDirection direction)
    {
        IReadOnlyList<ExcelMappingDiagnostic> diagnostics;
        return MigrateV1Json(json, direction, out diagnostics);
    }

    /// <summary>
    /// 将 v1 平铺 JSON 映射配置迁移到指定方向的 v2 文档，并返回诊断信息。
    /// </summary>
    /// <param name="json">v1 JSON 配置。</param>
    /// <param name="direction">迁移目标方向。</param>
    /// <param name="diagnostics">迁移诊断信息。</param>
    public static ExcelMappingDocument MigrateV1Json(string json, MappingDirection direction,
        out IReadOnlyList<ExcelMappingDiagnostic> diagnostics)
    {
        ValidateDirection(direction);
        var items = new List<ExcelMappingDiagnostic>();
        var configuration = DeserializeV1Json(json);
        var result = CreateMigratedDocument(configuration, direction);
        items.Add(new ExcelMappingDiagnostic("V1_MIGRATED", "$",
            $"检测到 v1 平铺 JSON，已显式迁移为 v2 {direction} 方向文档。"));
        diagnostics = items;
        return result;
    }

    /// <summary>
    /// 从调用方拥有的流迁移 v1 JSON，并保留调用方流所有权。
    /// </summary>
    /// <param name="source">v1 JSON 流。</param>
    /// <param name="direction">迁移目标方向。</param>
    public static ExcelMappingDocument MigrateV1Json(Stream source, MappingDirection direction)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("JSON 配置流不可读取。", nameof(source));
        using var reader = new StreamReader(source, Encoding.UTF8, true, 1024, true);
        return MigrateV1Json(ExcelMappingTextReader.ReadLimitedText(reader), direction);
    }
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
            ExcelMappingDocumentValidator.ValidateJsonElement(document.RootElement, "$", isV2);
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
                throw new InvalidOperationException(
                    "检测到 v1 平铺 JSON；请调用 MigrateV1Json(json, direction) 并显式指定迁移方向。");
            }
            ExcelMappingDocumentValidator.ValidateDocument(result, modelAliases);
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
        return FromJsonDocument(ExcelMappingTextReader.ReadLimitedText(reader), modelAliases);
    }

    /// <summary>
    /// 从 XML 文本加载规范化映射文档。
    /// </summary>
    public static ExcelMappingDocument FromXmlDocument(string xml)
        => LoadXmlDocument(xml, null, null);


    /// <summary>
    /// 将 v1 平铺 XML 映射配置迁移到指定方向的 v2 文档。
    /// </summary>
    /// <param name="xml">v1 XML 配置。</param>
    /// <param name="direction">迁移目标方向。</param>
    public static ExcelMappingDocument MigrateV1Xml(string xml, MappingDirection direction)
    {
        IReadOnlyList<ExcelMappingDiagnostic> diagnostics;
        return MigrateV1Xml(xml, direction, out diagnostics);
    }

    /// <summary>
    /// 将 v1 平铺 XML 映射配置迁移到指定方向的 v2 文档，并返回诊断信息。
    /// </summary>
    /// <param name="xml">v1 XML 配置。</param>
    /// <param name="direction">迁移目标方向。</param>
    /// <param name="diagnostics">迁移诊断信息。</param>
    public static ExcelMappingDocument MigrateV1Xml(string xml, MappingDirection direction,
        out IReadOnlyList<ExcelMappingDiagnostic> diagnostics)
    {
        ValidateDirection(direction);
        var items = new List<ExcelMappingDiagnostic>();
        var configuration = DeserializeV1Xml(xml);
        var result = CreateMigratedDocument(configuration, direction);
        items.Add(new ExcelMappingDiagnostic("V1_MIGRATED", "/ExcelMappingConfiguration",
            $"检测到 v1 平铺 XML，已显式迁移为 v2 {direction} 方向文档。"));
        diagnostics = items;
        return result;
    }

    /// <summary>
    /// 从调用方拥有的流迁移 v1 XML，并保留调用方流所有权。
    /// </summary>
    /// <param name="source">v1 XML 流。</param>
    /// <param name="direction">迁移目标方向。</param>
    public static ExcelMappingDocument MigrateV1Xml(Stream source, MappingDirection direction)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("XML 配置流不可读取。", nameof(source));
        using var reader = new StreamReader(source, Encoding.UTF8, true, 1024, true);
        return MigrateV1Xml(ExcelMappingTextReader.ReadLimitedText(reader), direction);
    }
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
        using (var shapeReader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings()))
        {
            var shape = XDocument.Load(shapeReader, LoadOptions.SetLineInfo);
            ExcelMappingDocumentValidator.ValidateXmlShape(shape.Root, isV2);
        }
        if (!isV2)
            throw new InvalidOperationException(
                "检测到 v1 平铺 XML；请调用 MigrateV1Xml(xml, direction) 并显式指定迁移方向。");
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        var result = DeserializeXml(reader);
        ExcelMappingDocumentValidator.ValidateDocument(result, modelAliases);
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
        return FromXmlDocument(ExcelMappingTextReader.ReadLimitedText(reader), modelAliases);
    }

    /// <summary>
    /// 将 normalized v2 文档写为 JSON。
    /// </summary>
    public static string ToJson(ExcelMappingDocument document)
    {
        ExcelMappingDocumentValidator.ValidateDocument(document, null);
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
        ExcelMappingDocumentValidator.ValidateDocument(document, null);
        var serializer = new XmlSerializer(typeof(ExcelMappingDocument));
        using var writer = new Utf8StringWriter();
        serializer.Serialize(writer, document);
        return writer.ToString();
    }

    private static ExcelMappingDocument CreateMigratedDocument(ExcelMappingConfiguration configuration,
        MappingDirection direction)
    {
        if (configuration == null)
            throw new InvalidOperationException("v1 配置未包含有效映射。");
        return new ExcelMappingDocument
        {
            Version = 2,
            Import = direction == MappingDirection.Import
                ? MappingConfigurationMerger.Merge(null, configuration, MappingSourceKind.Document)
                : new ExcelMappingConfiguration(),
            Export = direction == MappingDirection.Export
                ? MappingConfigurationMerger.Merge(null, configuration, MappingSourceKind.Document)
                : new ExcelMappingConfiguration()
        };
    }

    private static ExcelMappingConfiguration DeserializeV1Json(string json)
    {
        ExcelMappingTextReader.ValidateDocumentText(json, "JSON");
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
            ExcelMappingDocumentValidator.ValidateJsonElement(document.RootElement, "$", false);
            return JsonSerializer.Deserialize<ExcelMappingConfiguration>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                MaxDepth = MaxDepth
            }) ?? throw new InvalidOperationException("JSON 配置未包含有效映射。");
        }
        catch (JsonException exception)
        {
            throw new InvalidOperationException($"JSON 映射配置无效: {exception.Message}", exception);
        }
    }

    private static ExcelMappingConfiguration DeserializeV1Xml(string xml)
    {
        ExcelMappingTextReader.ValidateDocumentText(xml, "XML");
        using (var shapeReader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings()))
        {
            var shape = XDocument.Load(shapeReader, LoadOptions.SetLineInfo);
            ExcelMappingDocumentValidator.ValidateXmlShape(shape.Root, false);
        }
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        var serializer = new XmlSerializer(typeof(ExcelMappingConfiguration));
        AttachXmlValidationHandlers(serializer);
        return (ExcelMappingConfiguration)serializer.Deserialize(reader)
            ?? throw new InvalidOperationException("XML 配置未包含有效映射。");
    }

    private static void ValidateDirection(MappingDirection direction)
    {
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));
    }

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

    private static ExcelMappingDocument DeserializeXml(XmlReader reader)
    {
        try
        {
            reader.MoveToContent();
            var serializer = new XmlSerializer(typeof(ExcelMappingDocument));
            AttachXmlValidationHandlers(serializer);
            return (ExcelMappingDocument)serializer.Deserialize(reader)
                ?? throw new InvalidOperationException("XML 配置未包含有效映射文档。");
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
    public ExcelMappingDocument FromJsonDocument(string json) => ExcelMappingConfigurationLoader.FromJsonDocument(json);

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(Stream source) => ExcelMappingConfigurationLoader.FromJsonDocument(source);

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(string xml) => ExcelMappingConfigurationLoader.FromXmlDocument(xml);

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(Stream source) => ExcelMappingConfigurationLoader.FromXmlDocument(source);
}
