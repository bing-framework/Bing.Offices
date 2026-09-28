using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 已完成校验的不可变实体布局。
/// </summary>
/// <typeparam name="TEntity">聚合对象类型。</typeparam>
public sealed class ExcelEntityLayout<TEntity> where TEntity : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityLayout{TEntity}" /> 类型的实例。
    /// </summary>
    /// <param name="cells">固定单元格绑定快照。</param>
    /// <param name="listRegions">列表区域绑定快照。</param>
    /// <param name="merges">固定合并区域快照。</param>
    /// <param name="relations">父子关系绑定快照。</param>
    internal ExcelEntityLayout(IReadOnlyList<ExcelEntityCellBinding<TEntity>> cells,
        IReadOnlyList<ExcelEntityListRegion<TEntity>> listRegions,
        IReadOnlyList<ExcelEntityMergeRegion> merges,
        IReadOnlyList<ExcelRelationRequest> relations)
    {
        Cells = cells;
        ListRegions = listRegions;
        Merges = merges;
        Relations = Array.AsReadOnly((relations ?? Array.Empty<ExcelRelationRequest>()).ToArray());
    }

    /// <summary>
    /// 获取固定单元格绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityCellBinding<TEntity>> Cells { get; }

    /// <summary>
    /// 获取列表区域绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityListRegion<TEntity>> ListRegions { get; }

    /// <summary>
    /// 获取合并区域声明。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelEntityMergeRegion> Merges { get; }

    /// <summary>
    /// 获取聚合对象父子关系绑定。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelRelationRequest> Relations { get; }
}
