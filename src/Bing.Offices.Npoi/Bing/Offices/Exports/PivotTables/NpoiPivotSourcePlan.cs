using System.Collections.ObjectModel;
using Bing.Offices.Exports;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 保存已校验的透视来源及类型化筛选计划。
/// </summary>
internal sealed class NpoiPivotSourcePlan
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotSourcePlan" /> 类型的实例。
    /// </summary>
    /// <param name="definition">请求中的不可变透视定义。</param>
    /// <param name="sheet">实际来源工作表。</param>
    /// <param name="headers">完整来源表头。</param>
    /// <param name="rows">不含表头的全部字面值记录。</param>
    /// <param name="filters">按来源实际值解析的筛选选项。</param>
    /// <param name="outputRows">包括筛选区的完整展开行数上界。</param>
    /// <param name="outputColumns">完整展开列数上界。</param>
    internal NpoiPivotSourcePlan(ExcelPivotTableDefinition definition, XSSFSheet sheet,
        IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<object>> rows,
        IReadOnlyDictionary<string, IReadOnlyList<object>> filters, int outputRows, int outputColumns)
    {
        Definition = definition;
        Sheet = sheet;
        OutputRows = outputRows;
        OutputColumns = outputColumns;
        Headers = Array.AsReadOnly(headers.ToArray());
        Rows = Array.AsReadOnly(rows.Select(row => (IReadOnlyList<object>)
            Array.AsReadOnly(row.ToArray())).ToArray());
        Columns = new ReadOnlyDictionary<string, int>(Headers.Select((name, index) => (name, index))
            .ToDictionary(item => item.name, item => item.index, StringComparer.OrdinalIgnoreCase));
        Filters = new ReadOnlyDictionary<string, IReadOnlyList<object>>(filters.ToDictionary(
            pair => pair.Key, pair => (IReadOnlyList<object>)Array.AsReadOnly(pair.Value.ToArray()),
            StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// 获取请求中的透视定义。
    /// </summary>
    internal ExcelPivotTableDefinition Definition { get; }

    /// <summary>
    /// 获取实际来源工作表。
    /// </summary>
    internal XSSFSheet Sheet { get; }

    /// <summary>
    /// 获取完整来源表头。
    /// </summary>
    internal IReadOnlyList<string> Headers { get; }

    /// <summary>
    /// 获取完整来源记录快照。
    /// </summary>
    internal IReadOnlyList<IReadOnlyList<object>> Rows { get; }

    /// <summary>
    /// 获取字段到区域内零基列号的映射。
    /// </summary>
    internal IReadOnlyDictionary<string, int> Columns { get; }

    /// <summary>
    /// 获取按来源实际值解析的筛选选项。
    /// </summary>
    internal IReadOnlyDictionary<string, IReadOnlyList<object>> Filters { get; }

    /// <summary>
    /// 获取包含筛选区的完整展开行数上界。
    /// </summary>
    internal int OutputRows { get; }

    /// <summary>
    /// 获取完整展开列数上界。
    /// </summary>
    internal int OutputColumns { get; }
}
