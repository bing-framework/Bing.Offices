using System.Collections.ObjectModel;
using System.IO.Compression;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Internals;

/// <summary>
/// 读取透视模板部件与关系。
/// </summary>
/// <remarks>仅持有调用方的包引用，不接管其生命周期，也不加载工作表或缓存记录正文。</remarks>
internal sealed class ExcelPivotTemplatePackage
{
    /// <summary>
    /// 按实际部件名称保存 ZIP 条目。
    /// </summary>
    private readonly Dictionary<string, ZipArchiveEntry> _entries = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 按拥有者缓存只读关系集合。
    /// </summary>
    private readonly Dictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>> _relationships = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 读取关系和结构元数据时使用的 XML 限制。
    /// </summary>
    private readonly ExcelResourceLimits _limits;

    /// <summary>
    /// 当前模板预检所属的 Provider。
    /// </summary>
    private readonly string _provider;

    /// <summary>
    /// 当前模板预检所属的操作。
    /// </summary>
    private readonly BingOfficesOperation _operation;

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplatePackage" /> 类型的实例。
    /// </summary>
    /// <param name="archive">调用方持有的模板包。</param>
    /// <param name="token">取消令牌。</param>
    internal ExcelPivotTemplatePackage(ZipArchive archive, CancellationToken token)
        : this(archive, token, new ExcelResourceLimits
        {
            MaxXmlCharacters = null,
            MaxXmlDepth = null
        }, "Core", BingOfficesOperation.Import)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplatePackage" /> 类型的实例。
    /// </summary>
    /// <param name="archive">调用方持有的模板包。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="limits">关系和结构元数据的 XML 限制。</param>
    /// <param name="provider">当前模板预检所属的 Provider。</param>
    /// <param name="operation">当前模板预检所属的操作。</param>
    internal ExcelPivotTemplatePackage(ZipArchive archive, CancellationToken token,
        ExcelResourceLimits limits, string provider, BingOfficesOperation operation)
    {
        _limits = limits ?? throw new ArgumentNullException(nameof(limits));
        _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        _operation = operation;
        foreach (var entry in archive.Entries)
        {
            token.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                continue;
            var name = entry.FullName.Replace('\\', '/').TrimStart('/');
            if (!_entries.TryAdd(name, entry))
                throw new InvalidDataException($"模板包含重复部件: {name}");
        }
    }

    /// <summary>
    /// 扫描指定部件的 XML 元素。
    /// </summary>
    /// <param name="partName">实际部件名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="visit">元素处理委托。</param>
    internal void Read(string partName, CancellationToken token, Action<XmlReader> visit)
    {
        token.ThrowIfCancellationRequested();
        if (!_entries.TryGetValue(partName, out var entry))
            throw new InvalidDataException($"透视模板缺少部件: {partName}");
        ExcelPivotTemplateParts.ReadMetadata(entry, _limits, _provider, _operation, token, visit);
    }

    /// <summary>
    /// 读取并缓存部件拥有者的关系。
    /// </summary>
    /// <param name="owner">拥有者部件；空字符串表示包根。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>保持原始顺序的只读关系集合。</returns>
    internal IReadOnlyList<ExcelPivotTemplateRelationship> Relationships(string owner, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (_relationships.TryGetValue(owner, out var existing))
            return existing;
        var separator = owner.LastIndexOf('/');
        var relationshipPart = owner.Length == 0 ? "_rels/.rels"
            : owner[..(separator + 1)] + "_rels/" + owner[(separator + 1)..] + ".rels";
        var values = new List<ExcelPivotTemplateRelationship>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        if (_entries.ContainsKey(relationshipPart))
            Read(relationshipPart, token, reader =>
            {
                if (reader.LocalName != "Relationship"
                    || reader.NamespaceURI != "http://schemas.openxmlformats.org/package/2006/relationships")
                    return;
                var id = Required(reader.GetAttribute("Id"), relationshipPart);
                if (!ids.Add(id))
                    throw new InvalidDataException($"模板关系标识重复: {relationshipPart}/{id}");
                var type = Required(reader.GetAttribute("Type"), relationshipPart);
                var target = Required(reader.GetAttribute("Target"), relationshipPart);
                var external = string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                values.Add(new ExcelPivotTemplateRelationship(id, type, target,
                    external ? null : ExcelPivotTemplateParts.ResolveTarget(relationshipPart, target), external));
            });
        var snapshot = Array.AsReadOnly(values.ToArray());
        _relationships.Add(owner, snapshot);
        return snapshot;
    }

    /// <summary>
    /// 验证指定内部关系及目标部件。
    /// </summary>
    /// <param name="owner">关系拥有者。</param>
    /// <param name="id">关系标识。</param>
    /// <param name="kind">预期关系类型。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>有效关系。</returns>
    internal ExcelPivotTemplateRelationship RequireRelationship(string owner, string id,
        string kind, CancellationToken token)
    {
        var value = Relationships(owner, token).SingleOrDefault(item => item.Id == id);
        if (value == null || !IsRelationship(value.Type, kind) || value.External
            || !_entries.ContainsKey(value.PartName))
            throw new InvalidDataException($"透视模板关系无效: {owner}/{id}/{kind}");
        return value;
    }

    /// <summary>
    /// 获取已读取的全部关系快照。
    /// </summary>
    /// <returns>按拥有者索引的不可变集合。</returns>
    internal IReadOnlyDictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>> SnapshotRelationships() =>
        new ReadOnlyDictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>>(
            new Dictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>>(_relationships, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 判断过渡或严格格式的关系类型。
    /// </summary>
    /// <param name="type">实际关系类型。</param>
    /// <param name="kind">类型末段。</param>
    /// <returns>关系类型是否匹配。</returns>
    internal static bool IsRelationship(string type, string kind) =>
        type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + kind
        || type == "http://purl.oclc.org/ooxml/officeDocument/relationships/" + kind;

    /// <summary>
    /// 读取必填元数据。
    /// </summary>
    /// <param name="value">元数据值。</param>
    /// <param name="part">所属部件。</param>
    /// <returns>非空元数据。</returns>
    internal static string Required(string value, string part) =>
        string.IsNullOrWhiteSpace(value) ? throw new InvalidDataException($"透视模板元数据缺失: {part}") : value;
}
