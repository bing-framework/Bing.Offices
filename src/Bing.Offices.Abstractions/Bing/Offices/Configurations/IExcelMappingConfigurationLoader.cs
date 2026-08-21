namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 映射配置加载器。
/// </summary>
public interface IExcelMappingConfigurationLoader
{
    /// <summary>
    /// 从 JSON 文本加载映射配置。
    /// </summary>
    /// <param name="json">JSON 文本。</param>
    ExcelMappingConfiguration FromJson(string json);

    /// <summary>
    /// 从调用方拥有的 JSON 流加载映射配置。
    /// </summary>
    /// <param name="source">JSON 配置流。</param>
    ExcelMappingConfiguration FromJson(Stream source);

    /// <summary>
    /// 从 XML 文本加载映射配置。
    /// </summary>
    /// <param name="xml">XML 文本。</param>
    ExcelMappingConfiguration FromXml(string xml);

    /// <summary>
    /// 从调用方拥有的 XML 流加载映射配置。
    /// </summary>
    /// <param name="source">XML 配置流。</param>
    ExcelMappingConfiguration FromXml(Stream source);
}
