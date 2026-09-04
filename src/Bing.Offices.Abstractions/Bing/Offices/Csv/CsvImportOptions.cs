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
    /// 获取或设置规范化映射文档；导入器使用其 Import 方向。
    /// </summary>
    public ExcelMappingDocument MappingDocument { get; set; }

    /// <summary>
    /// 获取或设置唯一值跟踪上限；为空表示不额外限制。
    /// </summary>
    public int? MaxTrackedUniqueValues { get; set; }

    /// <summary>
    /// 获取或设置唯一值比较策略。
    /// </summary>
    public StringComparison UniqueComparison { get; set; } = StringComparison.OrdinalIgnoreCase;

    /// <summary>输入流最大字节数；为空表示不额外限制。</summary>
    public long? MaxInputBytes { get; set; }

    /// <summary>最大数据行数，不包含表头；为空表示不额外限制。</summary>
    public int? MaxRows { get; set; }

    /// <summary>最大错误数；为空表示不额外限制。</summary>
    public int? MaxErrors { get; set; }

    /// <summary>单字段最大字符数；为空表示不额外限制。</summary>
    public int? MaxFieldLength { get; set; }

    /// <summary>单条记录最大列数；为空表示不额外限制。</summary>
    public int? MaxColumns { get; set; }

    /// <summary>验证资源限制值。</summary>
    public void Validate()
    {
        if (MaxInputBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxRows <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxRows));
        if (MaxErrors <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxErrors));
        if (MaxFieldLength <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxFieldLength));
        if (MaxColumns <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxColumns));
        if (MaxTrackedUniqueValues <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTrackedUniqueValues));
        if (UniqueComparison != StringComparison.Ordinal && UniqueComparison != StringComparison.OrdinalIgnoreCase
            && UniqueComparison != StringComparison.InvariantCulture
            && UniqueComparison != StringComparison.InvariantCultureIgnoreCase
            && UniqueComparison != StringComparison.CurrentCulture
            && UniqueComparison != StringComparison.CurrentCultureIgnoreCase)
            throw new ArgumentOutOfRangeException(nameof(UniqueComparison));
    }
}
