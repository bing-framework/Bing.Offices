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
    /// <param name="lookupLists">文本字典快照。</param>
    /// <param name="cascadeLists">级联字典快照。</param>
    /// <param name="dataValidations">按工作表分组的原生校验快照。</param>
    /// <param name="protections">按工作表分组的保护定义。</param>
    /// <param name="printTitleRows">重复打印标题行。</param>
    internal ExcelEntityLayout(IReadOnlyList<ExcelEntityCellBinding<TEntity>> cells,
        IReadOnlyList<ExcelEntityListRegion<TEntity>> listRegions,
        IReadOnlyList<ExcelEntityMergeRegion> merges,
        IReadOnlyList<ExcelRelationRequest> relations,
        IReadOnlyList<ExcelLookupListDefinition> lookupLists,
        IReadOnlyList<ExcelCascadeListDefinition> cascadeLists,
        IReadOnlyDictionary<string, List<ExcelDataValidationDefinition>> dataValidations,
        IReadOnlyDictionary<string, ExcelSheetProtectionDefinition> protections,
        IReadOnlyDictionary<string, string> printTitleRows = null)
    {
        Cells = cells;
        ListRegions = listRegions;
        Merges = merges;
        Relations = Array.AsReadOnly((relations ?? Array.Empty<ExcelRelationRequest>()).ToArray());
        LookupLists = Array.AsReadOnly((lookupLists ?? Array.Empty<ExcelLookupListDefinition>()).ToArray());
        CascadeLists = Array.AsReadOnly((cascadeLists ?? Array.Empty<ExcelCascadeListDefinition>()).ToArray());
        DataValidations = new System.Collections.ObjectModel.ReadOnlyDictionary<string,
            IReadOnlyList<ExcelDataValidationDefinition>>(
            (dataValidations ?? new Dictionary<string, List<ExcelDataValidationDefinition>>())
            .ToDictionary(pair => pair.Key,
                pair => (IReadOnlyList<ExcelDataValidationDefinition>)Array.AsReadOnly(
                    pair.Value.Select(item => item.Clone()).ToArray()), StringComparer.OrdinalIgnoreCase));
        Protections = new System.Collections.ObjectModel.ReadOnlyDictionary<string,
            ExcelSheetProtectionDefinition>(
            (protections ?? new Dictionary<string, ExcelSheetProtectionDefinition>())
            .ToDictionary(pair => pair.Key, pair => pair.Value.Clone(), StringComparer.OrdinalIgnoreCase));
        PrintTitleRows = new System.Collections.ObjectModel.ReadOnlyDictionary<string, string>(
            (printTitleRows ?? new Dictionary<string, string>())
            .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.OrdinalIgnoreCase));
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

    /// <summary>
    /// 获取工作簿文本字典。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelLookupListDefinition> LookupLists { get; }

    /// <summary>
    /// 获取两级联动字典。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelCascadeListDefinition> CascadeLists { get; }

    /// <summary>
    /// 获取工作表原生校验定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyDictionary<string, IReadOnlyList<ExcelDataValidationDefinition>> DataValidations { get; }

    /// <summary>
    /// 获取工作表保护定义。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyDictionary<string, ExcelSheetProtectionDefinition> Protections { get; }

    /// <summary>获取各工作表的重复打印标题行。</summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyDictionary<string, string> PrintTitleRows { get; }
}
