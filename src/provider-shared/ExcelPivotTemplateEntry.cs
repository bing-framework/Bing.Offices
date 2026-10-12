namespace Bing.Offices.Internals;

/// <summary>
/// 已有透视表的所属工作表和缓存快照。
/// </summary>
internal sealed class ExcelPivotTemplateEntry
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateEntry" /> 类型的实例。
    /// </summary>
    /// <param name="name">原始透视名称。</param>
    /// <param name="sheetName">所属工作表名称。</param>
    /// <param name="sheetPart">所属工作表部件。</param>
    /// <param name="relationshipId">工作表引用透视的关系标识。</param>
    /// <param name="partName">透视定义部件。</param>
    /// <param name="targetReference">原始显示区域。</param>
    /// <param name="cache">完整缓存描述。</param>
    /// <param name="hasUnverifiedDependencies">是否包含当前更新路径无法证明安全的扩展或依赖。</param>
    internal ExcelPivotTemplateEntry(string name, string sheetName, string sheetPart,
        string relationshipId, string partName, string targetReference, ExcelPivotTemplateCache cache,
        bool hasUnverifiedDependencies)
    {
        Name = name;
        SheetName = sheetName;
        SheetPart = sheetPart;
        RelationshipId = relationshipId;
        PartName = partName;
        TargetReference = targetReference;
        Cache = cache;
        HasUnverifiedDependencies = hasUnverifiedDependencies;
    }

    /// <summary>
    /// 获取透视名称。
    /// </summary>
    internal string Name { get; }

    /// <summary>
    /// 获取所属工作表名称。
    /// </summary>
    internal string SheetName { get; }

    /// <summary>
    /// 获取工作表部件。
    /// </summary>
    internal string SheetPart { get; }

    /// <summary>
    /// 获取工作表关系标识。
    /// </summary>
    internal string RelationshipId { get; }

    /// <summary>
    /// 获取透视部件名称。
    /// </summary>
    internal string PartName { get; }

    /// <summary>
    /// 获取原始显示区域。
    /// </summary>
    internal string TargetReference { get; }

    /// <summary>
    /// 获取缓存描述。
    /// </summary>
    internal ExcelPivotTemplateCache Cache { get; }

    /// <summary>
    /// 获取是否包含未验证的扩展或依赖。
    /// </summary>
    internal bool HasUnverifiedDependencies { get; }
}
