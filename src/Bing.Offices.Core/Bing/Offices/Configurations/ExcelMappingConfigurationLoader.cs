using System.Text.Json;
using System.Xml;
using System.Xml.Serialization;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 映射配置加载器。
/// </summary>
public static class ExcelMappingConfigurationLoader
{
    /// <summary>
    /// 从 JSON 文本加载映射配置。
    /// </summary>
    /// <param name="json">JSON 文本。</param>
    public static ExcelMappingConfiguration FromJson(string json)
    {
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("JSON 配置不能为空。", nameof(json));
        return JsonSerializer.Deserialize<ExcelMappingConfiguration>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? throw new InvalidOperationException("JSON 配置未包含有效映射。");
    }

    /// <summary>
    /// 从 JSON 流加载映射配置。
    /// </summary>
    /// <param name="source">调用方拥有的 JSON 流。</param>
    public static ExcelMappingConfiguration FromJson(Stream source)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("JSON 配置流不可读取。", nameof(source));
        using var reader = new StreamReader(source, System.Text.Encoding.UTF8, true, 1024, true);
        return FromJson(reader.ReadToEnd());
    }

    /// <summary>
    /// 从 UTF-8 JSON 配置文件加载映射配置。
    /// </summary>
    /// <param name="path">配置文件路径。</param>
    public static ExcelMappingConfiguration FromJsonFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("JSON 配置文件路径不能为空。", nameof(path));
        return FromJson(File.ReadAllText(path, System.Text.Encoding.UTF8));
    }

    /// <summary>
    /// 从 XML 文本加载映射配置。
    /// </summary>
    /// <param name="xml">XML 文本。</param>
    public static ExcelMappingConfiguration FromXml(string xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
            throw new ArgumentException("XML 配置不能为空。", nameof(xml));
        using var reader = XmlReader.Create(new StringReader(xml), CreateXmlReaderSettings());
        return DeserializeXml(reader);
    }

    /// <summary>
    /// 从 XML 流加载映射配置。
    /// </summary>
    /// <param name="source">调用方拥有的 XML 流。</param>
    public static ExcelMappingConfiguration FromXml(Stream source)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("XML 配置流不可读取。", nameof(source));
        using var reader = XmlReader.Create(source, CreateXmlReaderSettings());
        return DeserializeXml(reader);
    }

    /// <summary>
    /// 从 UTF-8 XML 配置文件加载映射配置。
    /// </summary>
    /// <param name="path">配置文件路径。</param>
    public static ExcelMappingConfiguration FromXmlFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("XML 配置文件路径不能为空。", nameof(path));
        return FromXml(File.ReadAllText(path, System.Text.Encoding.UTF8));
    }

    /// <summary>
    /// 创建禁止 DTD 与外部实体的 XML 读取设置。
    /// </summary>
    private static XmlReaderSettings CreateXmlReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null
    };

    /// <summary>
    /// 反序列化 XML 映射配置。
    /// </summary>
    /// <param name="reader">安全 XML 读取器。</param>
    private static ExcelMappingConfiguration DeserializeXml(XmlReader reader)
    {
        try
        {
            return (ExcelMappingConfiguration)new XmlSerializer(typeof(ExcelMappingConfiguration)).Deserialize(reader)
                ?? throw new InvalidOperationException("XML 配置未包含有效映射。");
        }
        catch (InvalidOperationException exception) when (exception.InnerException is XmlException xmlException)
        {
            throw xmlException;
        }
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
    public ExcelMappingConfiguration FromXml(string xml) => ExcelMappingConfigurationLoader.FromXml(xml);

    /// <inheritdoc />
    public ExcelMappingConfiguration FromXml(Stream source) => ExcelMappingConfigurationLoader.FromXml(source);
}
