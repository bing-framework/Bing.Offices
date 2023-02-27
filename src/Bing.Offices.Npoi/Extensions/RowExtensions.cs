﻿using System;
using System.Linq;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 行(<see cref="IRow"/>) 扩展
/// </summary>
public static class RowExtensions
{
    /// <summary>
    /// 获取或创建单元格
    /// </summary>
    /// <param name="row">行</param>
    /// <param name="cellIndex">单元格索引</param>
    public static ICell GetOrCreateCell(this IRow row, int cellIndex) => row.GetCell(cellIndex) ?? row.CreateCell(cellIndex);

    /// <summary>
    /// 创建单元格并进行操作
    /// </summary>
    /// <param name="row">行</param>
    /// <param name="cellIndex">单元格索引</param>
    /// <param name="action">单元格操作</param>
    public static IRow CreateCell(this IRow row, int cellIndex, Action<ICell> action)
    {
        var cell = row.GetOrCreateCell(cellIndex);
        action?.Invoke(cell);
        return row;
    }

    /// <summary>
    /// 清空内容
    /// </summary>
    /// <param name="row">NPOI单元行</param>
    public static IRow ClearContent(this IRow row)
    {
        foreach (var cell in row.Cells)
            cell.SetCellValue(string.Empty);
        return row;
    }

    /// <summary>
    /// 是否空行
    /// </summary>
    /// <param name="row">行</param>
    public static bool IsEmptyRow(this IRow row) => row == null || row.Cells.All(x => string.IsNullOrWhiteSpace(x?.GetStringValue()));

    /// <summary>
    /// 设置单元格值
    /// </summary>
    /// <typeparam name="T">数据类型</typeparam>
    /// <param name="row">行</param>
    /// <param name="column">单元格索引</param>
    /// <param name="value">值</param>
    /// <param name="style">行样式</param>
    public static void Value<T>(this IRow row, int column, T value, ICellStyle style = null)
    {
        var cell = row.CreateCell(column);
        cell.SetValue(value);
        if (style != null)
            cell.CellStyle = style;
    }
}