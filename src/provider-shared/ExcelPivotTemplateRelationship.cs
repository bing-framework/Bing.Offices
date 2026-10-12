namespace Bing.Offices.Internals;

/// <summary>
/// 透视模板的原始包关系。
/// </summary>
internal sealed class ExcelPivotTemplateRelationship
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateRelationship" /> 类型的实例。
    /// </summary>
    /// <param name="id">关系标识。</param>
    /// <param name="type">关系类型。</param>
    /// <param name="target">原始关系目标。</param>
    /// <param name="partName">解析后的内部部件；外部关系为空。</param>
    /// <param name="external">是否为外部关系。</param>
    internal ExcelPivotTemplateRelationship(string id, string type, string target,
        string partName, bool external)
    {
        Id = id;
        Type = type;
        Target = target;
        PartName = partName;
        External = external;
    }

    /// <summary>
    /// 获取关系标识。
    /// </summary>
    internal string Id { get; }

    /// <summary>
    /// 获取关系类型。
    /// </summary>
    internal string Type { get; }

    /// <summary>
    /// 获取原始目标。
    /// </summary>
    internal string Target { get; }

    /// <summary>
    /// 获取内部部件名称。
    /// </summary>
    internal string PartName { get; }

    /// <summary>
    /// 获取是否为外部关系。
    /// </summary>
    internal bool External { get; }
}
