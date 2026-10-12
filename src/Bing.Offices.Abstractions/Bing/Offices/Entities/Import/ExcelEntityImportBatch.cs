using System;
using System.Collections.Generic;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体列表区域的一批导入结果。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityImportBatch<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityImportBatch{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="entity">完整实体导入得到的聚合对象。</param>
    /// <param name="sheetName">来源工作表名称；诊断批次无归属工作表时为空。</param>
    /// <param name="propertyName">列表区域绑定的实体集合属性名称；诊断批次为空。</param>
    /// <param name="itemType">列表项类型；诊断批次使用聚合对象类型。</param>
    /// <param name="regionIndex">布局列表区域索引；诊断批次为 -1。</param>
    /// <param name="batchNumber">从一开始的全局批次号。</param>
    /// <param name="firstRowIndex">批次第一行的一基物理行号；无行号时为 0。</param>
    /// <param name="items">本批成功实体。</param>
    /// <param name="errors">本批结构化错误。</param>
    /// <param name="isDiagnostic">是否为固定单元格、关系或其他实体级诊断批次。</param>
    public ExcelEntityImportBatch(TEntity entity, string sheetName, string propertyName, Type itemType,
        int regionIndex, int batchNumber, int firstRowIndex, IReadOnlyList<object> items,
        IReadOnlyList<Imports.ExcelImportError> errors, bool isDiagnostic)
    {
        Entity = entity ?? throw new ArgumentNullException(nameof(entity));
        SheetName = sheetName;
        PropertyName = propertyName;
        ItemType = itemType ?? throw new ArgumentNullException(nameof(itemType));
        if (regionIndex < -1)
            throw new ArgumentOutOfRangeException(nameof(regionIndex));
        if (batchNumber <= 0)
            throw new ArgumentOutOfRangeException(nameof(batchNumber));
        if (firstRowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(firstRowIndex));
        RegionIndex = regionIndex;
        BatchNumber = batchNumber;
        FirstRowIndex = firstRowIndex;
        Items = items ?? Array.Empty<object>();
        Errors = errors ?? Array.Empty<Imports.ExcelImportError>();
        IsDiagnostic = isDiagnostic;
    }

    /// <summary>
    /// 获取完整实体导入得到的聚合对象。
    /// </summary>
    public TEntity Entity { get; }

    /// <summary>
    /// 获取来源工作表名称。
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// 获取列表区域绑定的实体集合属性名称。
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// 获取本批列表项类型。
    /// </summary>
    public Type ItemType { get; }

    /// <summary>
    /// 获取布局中的列表区域索引；诊断批次为 -1。
    /// </summary>
    public int RegionIndex { get; }

    /// <summary>
    /// 获取从一开始的全局批次号。
    /// </summary>
    public int BatchNumber { get; }

    /// <summary>
    /// 获取批次第一行的一基物理行号；无行号时为 0。
    /// </summary>
    public int FirstRowIndex { get; }

    /// <summary>
    /// 获取本批成功实体。
    /// </summary>
    public IReadOnlyList<object> Items { get; }

    /// <summary>
    /// 获取本批结构化错误。
    /// </summary>
    public IReadOnlyList<Imports.ExcelImportError> Errors { get; }

    /// <summary>
    /// 获取当前批次是否只承载实体级诊断错误。
    /// </summary>
    public bool IsDiagnostic { get; }
}
