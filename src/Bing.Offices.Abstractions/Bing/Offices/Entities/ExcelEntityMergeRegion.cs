using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 固定合并区域描述。
/// </summary>
public sealed class ExcelEntityMergeRegion
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityMergeRegion" /> 类型的实例。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="range">合并区域的零基矩形坐标。</param>
    internal ExcelEntityMergeRegion(string sheetName, ExcelEntityCellRange range)
    {
        SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        Range = range;
    }

    /// <summary>
    /// 获取目标工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取待读取或写入的合并区域。
    /// </summary>
    public ExcelEntityCellRange Range { get; }
}
