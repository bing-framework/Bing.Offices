using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 映射配置加载器。
/// </summary>
public static class ExcelMappingConfigurationLoader
{
    /// <summary>
    /// JSON 或 XML 映射文档允许的最大 UTF-8 字节数（1 MiB）。
    /// </summary>
    private const int MaxDocumentBytes = ExcelMappingTextReader.MaxDocumentBytes;
    /// <summary>
    /// JSON 解析器和序列化器允许的最大嵌套深度（32 层）。
    /// </summary>
    private const int MaxDepth = 32;
    /// <summary>
    /// 单个映射方向允许声明的最大列数（1,000 列）。
    /// </summary>
    private const int MaxColumns = 1000;
    /// <summary>
    /// 单列允许声明的最大标题别名数量（100 个）。
    /// </summary>
    private const int MaxAliasesPerColumn = 100;
    /// <summary>
    /// 单列允许声明的最大校验规则数量（100 条）。
    /// </summary>
    private const int MaxValidationsPerColumn = 100;
    /// <summary>
    /// 映射配置中单个字符串字段允许的最大字符数（4,096 个）。
    /// </summary>
    private const int MaxStringLength = 4096;

    /// <summary>
    /// 从 JSON 文本加载规范化映射文档。
    /// </summary>
    /// <param name="json">待加载的 JSON 文本。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromJsonDocument(string json)
        => ExecuteConfiguration(() => LoadJsonDocument(json, null));

    /// <summary>
    /// 从 JSON 文本加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 提供别名注册表时校验模型别名。
    /// </remarks>
    /// <param name="json">待加载的 JSON 文本。</param>
    /// <param name="modelAliases">用于校验模型别名的注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromJsonDocument(string json, ExcelModelAliasRegistry modelAliases)
        => ExecuteConfiguration(() => LoadJsonDocument(json, modelAliases));

    /// <summary>
    /// 解析、验证并反序列化 v2 JSON 映射文档。
    /// </summary>
    /// <param name="json">待加载的 JSON 文本。</param>
    /// <param name="modelAliases">用于校验模型别名的可选注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    private static ExcelMappingDocument LoadJsonDocument(string json, ExcelModelAliasRegistry modelAliases)
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
            if (!isV2)
                throw new InvalidOperationException("仅支持 v2 ExcelMappingDocument JSON 配置。");
            ExcelMappingDocumentValidator.ValidateJsonElement(document.RootElement, "$", true);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                MaxDepth = MaxDepth
            };
            var result = JsonSerializer.Deserialize<ExcelMappingDocument>(json, options)
                ?? throw new InvalidOperationException("JSON 配置未包含有效映射文档。");
            if (result.Version != 2)
                throw new InvalidOperationException($"不支持的 JSON 映射文档版本: {result.Version}");
            ExcelMappingDocumentValidator.ValidateDocument(result, modelAliases);
            return result;
        }
        catch (JsonException exception)
        {
            throw new InvalidOperationException($"JSON 映射配置无效: {exception.Message}", exception);
        }
    }

    /// <summary>
    /// 从 JSON 流加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 读取完成后不关闭调用方提供的流。
    /// </remarks>
    /// <param name="source">待读取的 JSON 流。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromJsonDocument(Stream source)
        => FromJsonDocument(source, null);

    /// <summary>
    /// 从 JSON 流加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 读取完成后不关闭调用方提供的流；提供别名注册表时校验模型别名。
    /// </remarks>
    /// <param name="source">待读取的 JSON 流。</param>
    /// <param name="modelAliases">用于校验模型别名的注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
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
    /// <param name="xml">待加载的 XML 文本。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromXmlDocument(string xml)
        => ExecuteConfiguration(() => LoadXmlDocument(xml, null));

    /// <summary>
    /// 从 XML 文本加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 提供别名注册表时校验模型别名。
    /// </remarks>
    /// <param name="xml">待加载的 XML 文本。</param>
    /// <param name="modelAliases">用于校验模型别名的注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromXmlDocument(string xml, ExcelModelAliasRegistry modelAliases)
        => ExecuteConfiguration(() => LoadXmlDocument(xml, modelAliases));

    /// <summary>
    /// 在禁止 DTD 和外部解析器的设置下解析并验证 v2 XML 映射文档。
    /// </summary>
    /// <param name="xml">待加载的 XML 文本。</param>
    /// <param name="modelAliases">用于校验模型别名的可选注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    private static ExcelMappingDocument LoadXmlDocument(string xml, ExcelModelAliasRegistry modelAliases)
    {
        if (string.IsNullOrWhiteSpace(xml))
            throw new ArgumentException("XML 配置不能为空。", nameof(xml));
        if (Encoding.UTF8.GetByteCount(xml) > MaxDocumentBytes)
            throw new InvalidOperationException($"XML 配置超过最大字节数: {MaxDocumentBytes}");
        var isV2 = IsXmlDocumentRoot(xml);
        if (!isV2)
            throw new InvalidOperationException("仅支持 v2 ExcelMappingDocument XML 配置。");
        using (var shapeReader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings()))
        {
            var shape = XDocument.Load(shapeReader, LoadOptions.SetLineInfo);
            ExcelMappingDocumentValidator.ValidateXmlShape(shape.Root, true);
        }
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        var result = DeserializeXml(reader);
        ExcelMappingDocumentValidator.ValidateDocument(result, modelAliases);
        return result;
    }

    /// <summary>
    /// 从 XML 流加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 读取完成后不关闭调用方提供的流。
    /// </remarks>
    /// <param name="source">待读取的 XML 流。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
    public static ExcelMappingDocument FromXmlDocument(Stream source)
        => FromXmlDocument(source, null);

    /// <summary>
    /// 从 XML 流加载规范化映射文档。
    /// </summary>
    /// <remarks>
    /// 读取完成后不关闭调用方提供的流；提供别名注册表时校验模型别名。
    /// </remarks>
    /// <param name="source">待读取的 XML 流。</param>
    /// <param name="modelAliases">用于校验模型别名的注册表。</param>
    /// <returns>已通过结构和业务规则验证的映射文档。</returns>
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
    /// <param name="document">待验证并序列化的映射文档。</param>
    /// <returns>格式化后的 JSON 文本。</returns>
    public static string ToJson(ExcelMappingDocument document)
        => ExecuteConfiguration(() =>
        {
            ExcelMappingDocumentValidator.ValidateDocument(document, null);
            return JsonSerializer.Serialize(document, new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                IgnoreNullValues = true,
                WriteIndented = true
            });
        });

    /// <summary>
    /// 将 normalized v2 文档写为 XML。
    /// </summary>
    /// <param name="document">待验证并序列化的映射文档。</param>
    /// <returns>包含 UTF-8 声明的 XML 文本。</returns>
    public static string ToXml(ExcelMappingDocument document)
        => ExecuteConfiguration(() =>
        {
            ExcelMappingDocumentValidator.ValidateDocument(document, null);
            var serializer = new XmlSerializer(typeof(ExcelMappingDocument));
            using var writer = new Utf8StringWriter();
            serializer.Serialize(writer, document);
            return writer.ToString();
        });

    /// <summary>
    /// 执行配置加载或序列化操作，并统一包装未分类的配置异常。
    /// </summary>
    /// <typeparam name="T">操作的返回类型。</typeparam>
    /// <param name="action">待执行的配置操作。</param>
    /// <returns>配置操作产生的结果。</returns>
    private static T ExecuteConfiguration<T>(Func<T> action)
    {
        try
        {
            return action();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException("映射配置无效。", exception);
        }
    }

    /// <summary>
    /// 创建禁用 DTD、外部实体且限制文档规模的 XML 读取设置。
    /// </summary>
    /// <returns>用于不可信映射 XML 的安全读取设置。</returns>
    private static XmlReaderSettings CreateXmlReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        MaxCharactersInDocument = MaxDocumentBytes,
        MaxCharactersFromEntities = 0
    };

    /// <summary>
    /// 检查 XML 根元素是否为 v2 映射文档根节点。
    /// </summary>
    /// <param name="xml">待检查的 XML 文本。</param>
    /// <returns>根元素为 <see cref="ExcelMappingDocument"/> 时为 true。</returns>
    private static bool IsXmlDocumentRoot(string xml)
    {
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        reader.MoveToContent();
        return string.Equals(reader.LocalName, nameof(ExcelMappingDocument), StringComparison.Ordinal);
    }

    /// <summary>
    /// 反序列化 v2 XML 文档，并将序列化验证错误转换为调用方可读错误。
    /// </summary>
    /// <param name="reader">已定位到 XML 内容的安全读取器。</param>
    /// <returns>反序列化后的映射文档。</returns>
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

    /// <summary>
    /// 从异常链中查找指定类型的首个异常。
    /// </summary>
    /// <typeparam name="TException">要查找的异常类型。</typeparam>
    /// <param name="exception">异常链起点。</param>
    /// <returns>匹配的内部异常；未找到时为 null。</returns>
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

    /// <summary>
    /// 为 XML 序列化器注册未知节点和属性的拒绝处理器。
    /// </summary>
    /// <param name="serializer">要配置的 XML 序列化器。</param>
    private static void AttachXmlValidationHandlers(XmlSerializer serializer)
    {
        serializer.UnknownNode += (_, eventArgs) =>
        throw new XmlMappingValidationException($"未知 XML 字段: /ExcelMappingDocument/{eventArgs.Name}");
        serializer.UnknownAttribute += (_, eventArgs) =>
        throw new XmlMappingValidationException($"未知 XML 属性: /ExcelMappingDocument/@{eventArgs.Attr?.Name ?? eventArgs.Attr?.LocalName}");
    }

    /// <summary>
    /// 表示映射 XML 包含未支持节点或属性的异常。
    /// </summary>
    private sealed class XmlMappingValidationException : InvalidOperationException
    {
        /// <summary>
        /// 初始化一个 <see cref="XmlMappingValidationException" /> 类型的实例。
        /// </summary>
        /// <param name="message">描述未知或无效 XML 成员的消息。</param>
        public XmlMappingValidationException(string message) : base(message)
        {
        }
    }

    /// <summary>
    /// 以 UTF-8 声明编码的字符串写入器。
    /// </summary>
    private sealed class Utf8StringWriter : StringWriter
    {
        /// <inheritdoc />
        /// <remarks>始终返回 UTF-8，以确保序列化 XML 声明与实际输出编码一致。</remarks>
        public override Encoding Encoding => Encoding.UTF8;
    }
}

/// <summary>
/// 映射配置加载器的默认服务实现。
/// </summary>
internal sealed class DefaultExcelMappingConfigurationLoader : IExcelMappingConfigurationLoader
{
    /// <summary>
    /// 向注册的观察器转发配置加载异常。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个 <see cref="DefaultExcelMappingConfigurationLoader" /> 类型的实例。
    /// </summary>
    /// <param name="exceptionObservers">接收配置加载异常的可选观察器集合。</param>
    public DefaultExcelMappingConfigurationLoader(IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(string json) =>
        Execute(() => ExcelMappingConfigurationLoader.FromJsonDocument(json));

    /// <inheritdoc />
    public ExcelMappingDocument FromJsonDocument(Stream source) =>
        Execute(() => ExcelMappingConfigurationLoader.FromJsonDocument(source));

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(string xml) =>
        Execute(() => ExcelMappingConfigurationLoader.FromXmlDocument(xml));

    /// <inheritdoc />
    public ExcelMappingDocument FromXmlDocument(Stream source) =>
        Execute(() => ExcelMappingConfigurationLoader.FromXmlDocument(source));

    /// <summary>
    /// 执行默认加载操作，并将配置异常通知观察器。
    /// </summary>
    /// <param name="load">待执行的映射文档加载操作。</param>
    /// <returns>加载后的映射文档。</returns>
    private ExcelMappingDocument Execute(Func<ExcelMappingDocument> load)
    {
        try
        {
            return load();
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }
}
