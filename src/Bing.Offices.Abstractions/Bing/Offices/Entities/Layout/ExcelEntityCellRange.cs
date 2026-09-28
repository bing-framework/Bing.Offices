using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Entities.Layout;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 零基 Excel 矩形区域。
/// </summary>
public readonly struct ExcelEntityCellRange : IEquatable<ExcelEntityCellRange>
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellRange" /> 类型的实例。
    /// </summary>
    /// <param name="first">区域左上角坐标。</param>
    /// <param name="last">区域右下角坐标。</param>
    internal ExcelEntityCellRange(ExcelEntityCellReference first, ExcelEntityCellReference last)
    {
        if (last.Row < first.Row || last.Column < first.Column)
            throw new ArgumentException("区域结束坐标不能位于起始坐标之前。");
        First = first;
        Last = last;
    }

    /// <summary>
    /// 获取区域的左上角坐标。
    /// </summary>
    public ExcelEntityCellReference First { get; }
    /// <summary>
    /// 获取区域的右下角坐标。
    /// </summary>
    public ExcelEntityCellReference Last { get; }
    /// <summary>
    /// 获取与当前零基坐标对应的 A1 区域地址。
    /// </summary>
    public string Address => First.Equals(Last) ? First.Address : $"{First.Address}:{Last.Address}";
    /// <summary>
    /// 判断指定坐标是否位于当前区域内。
    /// </summary>
    /// <param name="reference">要检查的零基单元格坐标。</param>
    /// <returns>坐标位于区域内时返回 <see langword="true" />；否则返回 <see langword="false" />。</returns>
    public bool Contains(ExcelEntityCellReference reference) => reference.Row >= First.Row && reference.Row <= Last.Row
        && reference.Column >= First.Column && reference.Column <= Last.Column;
    /// <summary>
    /// 解析 A1 矩形区域地址。
    /// </summary>
    /// <param name="range">要解析的单元格或矩形区域地址。</param>
    /// <returns>解析得到的零基矩形区域。</returns>
    public static ExcelEntityCellRange Parse(string range) => ExcelEntityAddress.ParseRange(range);
    /// <inheritdoc />
    public bool Equals(ExcelEntityCellRange other) => First.Equals(other.First) && Last.Equals(other.Last);
    /// <inheritdoc />
    public override bool Equals(object obj) => obj is ExcelEntityCellRange other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked((First.GetHashCode() * 397) ^ Last.GetHashCode());
    /// <inheritdoc />
    public override string ToString() => Address;
}
