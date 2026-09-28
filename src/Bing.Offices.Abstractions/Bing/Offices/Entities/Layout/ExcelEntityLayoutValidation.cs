using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities.Layout;

/// <summary>
/// 提供实体布局名称、包含关系和区域重叠的验证操作。
/// </summary>
internal static class ExcelEntityLayoutValidation
{
    /// <summary>
    /// 验证工作表名称是否满足 Excel 长度和字符约束。
    /// </summary>
    /// <param name="name">待验证的工作表名称。</param>
    internal static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称无效: {name}", nameof(name));
    }

    /// <summary>
    /// 验证模板命名锚点名称。
    /// </summary>
    /// <param name="name">待验证的命名锚点名称。</param>
    internal static void ValidateAnchorName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("模板命名锚点名称不能为空。", nameof(name));
        if (name.Length > 255 || name.IndexOfAny(new[] { '!', '[', ']', '\'', ' ' }) >= 0)
            throw new ArgumentException($"模板命名锚点名称无效: {name}", nameof(name));
    }

    /// <summary>
    /// 判断单元格是否位于指定起始坐标和可选结束坐标内。
    /// </summary>
    /// <param name="start">区域起始坐标。</param>
    /// <param name="end">可选的区域右下角坐标。</param>
    /// <param name="cell">待检查的单元格坐标。</param>
    /// <returns>单元格位于区域内时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Contains(ExcelEntityCellReference start, ExcelEntityCellReference? end,
        ExcelEntityCellReference cell) => cell.Row >= start.Row && cell.Column >= start.Column
        && (!end.HasValue || (cell.Row <= end.Value.Row && cell.Column <= end.Value.Column));

    /// <summary>
    /// 判断结束坐标是否位于起始坐标的右下方向。
    /// </summary>
    /// <param name="start">区域起始坐标。</param>
    /// <param name="end">区域结束坐标。</param>
    /// <returns>结束坐标不早于起始坐标时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Contains(ExcelEntityCellReference start, ExcelEntityCellReference end) =>
        end.Row >= start.Row && end.Column >= start.Column;

    /// <summary>
    /// 判断两个有界矩形区域是否存在交集。
    /// </summary>
    /// <param name="left">第一个矩形区域。</param>
    /// <param name="right">第二个矩形区域。</param>
    /// <returns>区域存在交集时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps(ExcelEntityCellRange left, ExcelEntityCellRange right) =>
        left.First.Row <= right.Last.Row && right.First.Row <= left.Last.Row
        && left.First.Column <= right.Last.Column && right.First.Column <= left.Last.Column;

    /// <summary>
    /// 判断两个列表区域是否存在可确定的交集。
    /// </summary>
    /// <typeparam name="TEntity">列表区域所属的聚合对象类型。</typeparam>
    /// <param name="left">第一个列表区域。</param>
    /// <param name="right">第二个列表区域。</param>
    /// <returns>已知边界或起点表明区域重叠时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps<TEntity>(ExcelEntityListRegion<TEntity> left,
        ExcelEntityListRegion<TEntity> right) where TEntity : class, new()
    {
        if (!left.End.HasValue && !right.End.HasValue)
            return true;
        if (left.End.HasValue && right.End.HasValue)
            return Overlaps(new ExcelEntityCellRange(left.Start, left.End.Value),
                new ExcelEntityCellRange(right.Start, right.End.Value));
        if (left.Start.Equals(right.Start))
            return true;
        return left.End.HasValue
            ? Contains(left.Start, left.End, right.Start)
            : right.End.HasValue && Contains(right.Start, right.End, left.Start);
    }

    /// <summary>
    /// 判断列表区域与固定矩形区域是否存在可确定的交集。
    /// </summary>
    /// <typeparam name="TEntity">列表区域所属的聚合对象类型。</typeparam>
    /// <param name="region">待检查的列表区域。</param>
    /// <param name="range">待检查的固定矩形区域。</param>
    /// <returns>区域存在可确定交集时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    internal static bool Overlaps<TEntity>(ExcelEntityListRegion<TEntity> region,
        ExcelEntityCellRange range) where TEntity : class, new()
    {
        if (region.End.HasValue)
            return Overlaps(new ExcelEntityCellRange(region.Start, region.End.Value), range);
        return region.Start.Row <= range.Last.Row && region.Start.Column <= range.Last.Column;
    }
}
