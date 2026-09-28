using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 冻结窗格定义。
/// </summary>
public sealed class ExcelFreezePaneDefinition
{
    /// <summary>
    /// 获取或初始化冻结的行数。
    /// </summary>
    public int Rows { get; init; }
    /// <summary>
    /// 获取或初始化冻结的列数。
    /// </summary>
    public int Columns { get; init; }
    /// <summary>
    /// 获取或初始化可选的顶部可视行索引。
    /// </summary>
    public int? TopRow { get; init; }
    /// <summary>
    /// 获取或初始化可选的左侧可视列索引。
    /// </summary>
    public int? LeftColumn { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (Rows < 0 || Columns < 0 || (Rows == 0 && Columns == 0))
            throw new ArgumentException("冻结窗格至少需要冻结一行或一列。", parameterName ?? nameof(Rows));
        if (TopRow.HasValue && TopRow.Value < 0 || LeftColumn.HasValue && LeftColumn.Value < 0)
            throw new ArgumentException("冻结窗格可视起点不能为负数。", parameterName ?? nameof(TopRow));
    }
}
