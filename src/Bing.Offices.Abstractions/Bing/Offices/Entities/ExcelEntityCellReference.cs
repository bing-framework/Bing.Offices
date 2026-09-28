using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 零基 Excel 单元格坐标。
/// </summary>
public readonly struct ExcelEntityCellReference : IEquatable<ExcelEntityCellReference>
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityCellReference" /> 类型的实例。
    /// </summary>
    /// <param name="row">从零开始的行索引。</param>
    /// <param name="column">从零开始的列索引。</param>
    internal ExcelEntityCellReference(int row, int column)
    {
        if (row < 0)
            throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0)
            throw new ArgumentOutOfRangeException(nameof(column));
        Row = row;
        Column = column;
    }

    /// <summary>
    /// 获取从零开始的行索引。
    /// </summary>
    public int Row { get; }
    /// <summary>
    /// 获取从零开始的列索引。
    /// </summary>
    public int Column { get; }
    /// <summary>
    /// 获取与当前零基坐标对应的 A1 地址。
    /// </summary>
    public string Address => ExcelEntityAddress.Format(Row, Column);

    /// <summary>
    /// 解析单个 A1 单元格地址。
    /// </summary>
    /// <param name="address">要解析的 A1 地址。</param>
    /// <returns>解析得到的零基单元格坐标。</returns>
    public static ExcelEntityCellReference Parse(string address) => ExcelEntityAddress.Parse(address);
    /// <inheritdoc />
    public bool Equals(ExcelEntityCellReference other) => Row == other.Row && Column == other.Column;
    /// <inheritdoc />
    public override bool Equals(object obj) => obj is ExcelEntityCellReference other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked((Row * 397) ^ Column);
    /// <inheritdoc />
    public override string ToString() => Address;
}
