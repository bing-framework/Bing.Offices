namespace Bing.Offices.Internals;

/// <summary>
/// 透视模板的缓存与消费者快照。
/// </summary>
internal sealed class ExcelPivotTemplateCache
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateCache" /> 类型的实例。
    /// </summary>
    /// <param name="id">工作簿缓存标识。</param>
    /// <param name="relationshipId">工作簿缓存关系标识。</param>
    /// <param name="definitionPart">缓存定义部件。</param>
    /// <param name="recordsPart">可选缓存记录部件。</param>
    /// <param name="recordsRelationshipId">可选记录关系标识。</param>
    /// <param name="sourceType">原始来源类型。</param>
    /// <param name="sourceSheet">来源工作表。</param>
    /// <param name="sourceReference">来源区域。</param>
    /// <param name="sourceName">可选命名来源。</param>
    /// <param name="fields">按索引排列的字段名称。</param>
    /// <param name="consumers">引用本缓存的透视部件。</param>
    /// <param name="hasUnverifiedDependencies">是否包含当前更新路径无法证明安全的扩展或依赖。</param>
    internal ExcelPivotTemplateCache(uint id, string relationshipId, string definitionPart,
        string recordsPart, string recordsRelationshipId, string sourceType, string sourceSheet,
        string sourceReference, string sourceName, IEnumerable<string> fields, IEnumerable<string> consumers,
        bool hasUnverifiedDependencies)
    {
        Id = id;
        RelationshipId = relationshipId;
        DefinitionPart = definitionPart;
        RecordsPart = recordsPart;
        RecordsRelationshipId = recordsRelationshipId;
        SourceType = sourceType;
        SourceSheet = sourceSheet;
        SourceReference = sourceReference;
        SourceName = sourceName;
        Fields = Array.AsReadOnly(fields.ToArray());
        Consumers = Array.AsReadOnly(consumers.ToArray());
        HasUnverifiedDependencies = hasUnverifiedDependencies;
    }

    /// <summary>
    /// 获取缓存标识。
    /// </summary>
    internal uint Id { get; }

    /// <summary>
    /// 获取工作簿关系标识。
    /// </summary>
    internal string RelationshipId { get; }

    /// <summary>
    /// 获取缓存定义部件。
    /// </summary>
    internal string DefinitionPart { get; }

    /// <summary>
    /// 获取缓存记录部件。
    /// </summary>
    internal string RecordsPart { get; }

    /// <summary>
    /// 获取记录关系标识。
    /// </summary>
    internal string RecordsRelationshipId { get; }

    /// <summary>
    /// 获取来源类型。
    /// </summary>
    internal string SourceType { get; }

    /// <summary>
    /// 获取来源工作表。
    /// </summary>
    internal string SourceSheet { get; }

    /// <summary>
    /// 获取来源区域。
    /// </summary>
    internal string SourceReference { get; }

    /// <summary>
    /// 获取命名来源。
    /// </summary>
    internal string SourceName { get; }

    /// <summary>
    /// 获取字段名称。
    /// </summary>
    internal IReadOnlyList<string> Fields { get; }

    /// <summary>
    /// 获取消费者部件。
    /// </summary>
    internal IReadOnlyList<string> Consumers { get; }

    /// <summary>
    /// 获取是否包含未验证的扩展或依赖。
    /// </summary>
    internal bool HasUnverifiedDependencies { get; }
}
