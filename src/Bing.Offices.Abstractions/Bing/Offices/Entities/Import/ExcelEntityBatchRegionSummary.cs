using System;

namespace Bing.Offices.Entities;

/// <summary>
/// 单个实体列表区域的分批导入摘要。
/// </summary>
public sealed class ExcelEntityBatchRegionSummary
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityBatchRegionSummary" /> 类型的实例。
    /// </summary>
    /// <param name="regionIndex">布局中的列表区域索引。</param>
    /// <param name="sheetName">来源工作表名称。</param>
    /// <param name="propertyName">列表区域绑定的实体集合属性名称。</param>
    /// <param name="itemType">列表项类型。</param>
    /// <param name="summary">区域批次统计。</param>
    public ExcelEntityBatchRegionSummary(int regionIndex, string sheetName, string propertyName,
        Type itemType, Imports.ExcelBatchImportSummary summary)
    {
        if (regionIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(regionIndex));
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(sheetName));
        if (string.IsNullOrWhiteSpace(propertyName))
            throw new ArgumentException("列表区域属性名称不能为空。", nameof(propertyName));
        RegionIndex = regionIndex;
        SheetName = sheetName;
        PropertyName = propertyName;
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        Summary = summary ?? throw new ArgumentNullException(nameof(summary));
    }

    /// <summary>
    /// 获取布局中的列表区域索引。
    /// </summary>
    public int RegionIndex { get; }

    /// <summary>
    /// 获取来源工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取列表区域绑定的实体集合属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取列表项类型。
    /// </summary>
    public Type ItemType { get; }

    /// <summary>
    /// 获取区域批次统计。
    /// </summary>
    public Imports.ExcelBatchImportSummary Summary { get; }
}
