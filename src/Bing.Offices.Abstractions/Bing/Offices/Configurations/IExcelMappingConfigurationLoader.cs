namespace Bing.Offices.Configurations;

/// <summary>
/// 加载并校验 JSON 或 XML 格式的 Excel v2 映射配置。
/// </summary>
public interface IExcelMappingConfigurationLoader
{
    /// <summary>
    /// 从 JSON 文本加载 v2 规范化映射文档。
    /// </summary>
    /// <param name="json">JSON 文本。</param>
    /// <returns>已通过结构和业务规则验证的规范化映射文档。</returns>
    ExcelMappingDocument FromJsonDocument(string json);

    /// <summary>
    /// 从调用方拥有的 JSON 流加载 v2 规范化映射文档。
    /// </summary>
    /// <param name="source">调用方拥有且可读取的 JSON 配置流；读取完成后保持打开。</param>
    /// <returns>已通过结构和业务规则验证的规范化映射文档。</returns>
    ExcelMappingDocument FromJsonDocument(Stream source);

    /// <summary>
    /// 从 XML 文本加载 v2 规范化映射文档。
    /// </summary>
    /// <param name="xml">XML 文本。</param>
    /// <returns>已通过结构和业务规则验证的规范化映射文档。</returns>
    ExcelMappingDocument FromXmlDocument(string xml);

    /// <summary>
    /// 从调用方拥有的 XML 流加载 v2 规范化映射文档。
    /// </summary>
    /// <param name="source">调用方拥有且可读取的 XML 配置流；读取完成后保持打开。</param>
    /// <returns>已通过结构和业务规则验证的规范化映射文档。</returns>
    ExcelMappingDocument FromXmlDocument(Stream source);
}
