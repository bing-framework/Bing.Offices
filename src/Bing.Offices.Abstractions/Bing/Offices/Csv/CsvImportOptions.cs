using System.Text;
using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Csv;

/// <summary>
/// CSV 流式导入选项。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
public sealed class CsvImportOptions<T> where T : class, new()
{
    /// <summary>
    /// 获取或设置是否包含表头。
    /// </summary>
    public bool HasHeader { get; set; } = true;

    /// <summary>
    /// 获取或设置字段分隔符。
    /// </summary>
    public char Delimiter { get; set; } = ',';

    /// <summary>
    /// 获取或设置字段引用字符。
    /// </summary>
    public char Quote { get; set; } = '"';

    /// <summary>
    /// 获取或设置值转换使用的区域性。
    /// </summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>
    /// 获取或设置文本编码。
    /// </summary>
    public Encoding Encoding { get; set; } = new UTF8Encoding(false);

    /// <summary>
    /// 获取或设置是否要求固定属性表头完整匹配。
    /// </summary>
    public bool HeaderMatch { get; set; } = true;

    /// <summary>
    /// 获取或设置本次导入的请求级映射配置。
    /// </summary>
    public ExcelMappingConfiguration MappingConfiguration { get; set; }

    /// <summary>
    /// 获取或设置本次导入使用的 Fluent 映射 Profile。
    /// </summary>
    public ExcelMappingProfile<T> MappingProfile { get; set; }
}
