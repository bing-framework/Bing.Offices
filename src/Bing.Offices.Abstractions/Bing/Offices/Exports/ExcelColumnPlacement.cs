using System;
using System.Collections.Generic;

namespace Bing.Offices.Exports;

/// <summary>
/// 动态列相对于固定列的请求级位置。
/// </summary>
public sealed class ExcelColumnPlacement
{
    private ExcelColumnPlacement(string beforeKey, string afterKey, int? physicalColumnIndex)
    {
        BeforeKey = beforeKey;
        AfterKey = afterKey;
        PhysicalColumnIndex = physicalColumnIndex;
    }

    /// <summary>
    /// 获取位于该固定列之前的列键。
    /// </summary>
    public string BeforeKey { get; }

    /// <summary>
    /// 获取位于该固定列之后的列键。
    /// </summary>
    public string AfterKey { get; }

    /// <summary>
    /// 获取显式物理列索引。
    /// </summary>
    public int? PhysicalColumnIndex { get; }

    /// <summary>
    /// 创建相对于固定列之前的位置。
    /// </summary>
    public static ExcelColumnPlacement Before(string key) => Create(key, null, null);

    /// <summary>
    /// 创建相对于固定列之后的位置。
    /// </summary>
    public static ExcelColumnPlacement After(string key) => Create(null, key, null);

    /// <summary>
    /// 创建显式物理列索引位置。
    /// </summary>
    public static ExcelColumnPlacement At(int columnIndex) => Create(null, null, columnIndex);

    internal static ExcelColumnPlacement Create(string beforeKey, string afterKey, int? physicalColumnIndex)
    {
        if (physicalColumnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(physicalColumnIndex));
        if (beforeKey != null && afterKey != null)
            throw new ArgumentException("动态列位置不能同时指定 Before 和 After。", nameof(beforeKey));
        if (beforeKey == null && afterKey == null && physicalColumnIndex == null)
            throw new ArgumentException("动态列位置必须指定 Before、After 或物理列索引。", nameof(beforeKey));
        if ((beforeKey != null || afterKey != null) && physicalColumnIndex != null)
            throw new ArgumentException("动态列位置不能同时指定相对位置和物理列索引。", nameof(physicalColumnIndex));
        return new ExcelColumnPlacement(beforeKey, afterKey, physicalColumnIndex);
    }
}
