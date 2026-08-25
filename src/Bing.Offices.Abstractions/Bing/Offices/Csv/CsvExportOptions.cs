using System.Text;
using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Csv;

/// <summary>
/// CSV 流式导出选项。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
public sealed class CsvExportOptions<T> where T : class, new()
{
    /// <summary>
    /// 获取或设置是否写入表头。
    /// </summary>
    public bool IncludeHeader { get; set; } = true;

    /// <summary>
    /// 获取或设置字段分隔符。
    /// </summary>
    public char Delimiter { get; set; } = ',';

    /// <summary>
    /// 获取或设置字段引用字符。
    /// </summary>
    public char Quote { get; set; } = '"';

    /// <summary>
    /// 获取或设置记录换行符。
    /// </summary>
    public string NewLine { get; set; } = "\r\n";

    /// <summary>
    /// 获取或设置值格式化和转换使用的区域性。
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// 获取或设置文本编码。
    /// </summary>
    public Encoding Encoding { get; set; } = new UTF8Encoding(false);

    /// <summary>
    /// 获取或设置潜在公式字段的处理策略。
    /// </summary>
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy { get; set; } = CsvFormulaInjectionPolicy.Escape;

    /// <summary>
    /// 获取或设置本次导出的请求级映射配置。
    /// </summary>
    public ExcelMappingConfiguration MappingConfiguration { get; set; }

    /// <summary>
    /// 获取或设置规范化映射文档；导出器使用其 Export 方向。
    /// </summary>
    public ExcelMappingDocument MappingDocument { get; set; }

    /// <summary>
    /// 获取或设置动态列名称。
    /// </summary>
    public IReadOnlyList<string> DynamicColumns { get; set; } = Array.Empty<string>();
}
