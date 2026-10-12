using System.IO.Compression;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Internals;

/// <summary>
/// 发现透视模板的实际 OPC 部件。
/// </summary>
internal static class ExcelPivotTemplateParts
{
    /// <summary>
    /// 内容类型描述的 XML 命名空间。
    /// </summary>
    private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

    /// <summary>
    /// 包关系描述的 XML 命名空间。
    /// </summary>
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// 透视内容类型前缀。
    /// </summary>
    private const string ContentTypePrefix = "application/vnd.openxmlformats-officedocument.spreadsheetml.";

    /// <summary>
    /// 过渡格式的文档关系前缀。
    /// </summary>
    private const string TransitionalRelationshipPrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

    /// <summary>
    /// 严格格式的文档关系前缀。
    /// </summary>
    private const string StrictRelationshipPrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    /// <summary>
    /// 判断部件是否属于透视表或透视缓存。
    /// </summary>
    /// <param name="partName">包内的部件名称。</param>
    /// <param name="contentType">部件内容类型。</param>
    /// <returns>是否属于透视部件。</returns>
    internal static bool IsPivotPart(string partName, string contentType)
    {
        var name = partName?.Replace('\\', '/').TrimStart('/') ?? string.Empty;
        return name.StartsWith("xl/pivotTables/", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("xl/pivotCache", StringComparison.OrdinalIgnoreCase)
            || IsPivotContentType(contentType);
    }

    /// <summary>
    /// 查找包内的透视部件。
    /// </summary>
    /// <param name="archive">由调用方持有的模板包。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按名称稳定排序的实际部件清单。</returns>
    /// <remarks>同时检查内容类型和关系目标，标准目录仅作为兼容发现依据。</remarks>
    internal static IReadOnlyList<string> Find(ZipArchive archive, CancellationToken cancellationToken = default)
        => Find(archive, new ExcelResourceLimits
        {
            MaxXmlCharacters = null,
            MaxXmlDepth = null
        }, "Core", BingOfficesOperation.Import, cancellationToken);

    /// <summary>
    /// 查找包内的透视部件。
    /// </summary>
    /// <param name="archive">由调用方持有的模板包。</param>
    /// <param name="limits">XML 字符数和嵌套深度限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按名称稳定排序的实际部件清单。</returns>
    /// <remarks>同时检查内容类型和关系目标，标准目录仅作为兼容发现依据。</remarks>
    /// <exception cref="BingOfficesResourceLimitException">XML 字符数或嵌套深度超过限制。</exception>
    internal static IReadOnlyList<string> Find(ZipArchive archive, ExcelResourceLimits limits,
        string provider, BingOfficesOperation operation, CancellationToken cancellationToken = default)
    {
        if (archive == null)
            throw new ArgumentNullException(nameof(archive));
        if (limits == null)
            throw new ArgumentNullException(nameof(limits));
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        var found = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in archive.Entries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                continue;
            var name = entry.FullName.Replace('\\', '/').TrimStart('/');
            if (!entries.TryAdd(name, entry))
                throw new InvalidDataException($"模板包含重复部件: {name}");
            if (IsPivotPart(name, null))
                found.Add(name);
        }

        if (entries.TryGetValue("[Content_Types].xml", out var contentTypes))
        {
            ReadMetadata(contentTypes, limits, provider, operation, cancellationToken, reader =>
            {
                if (reader.NamespaceURI != ContentTypesNamespace)
                    return;
                var contentType = reader.GetAttribute("ContentType");
                if (!IsPivotContentType(contentType))
                    return;
                if (reader.LocalName == "Override")
                    AddTarget(reader.GetAttribute("PartName")?.TrimStart('/'));
                else if (reader.LocalName == "Default")
                {
                    var extension = reader.GetAttribute("Extension");
                    foreach (var name in entries.Keys)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (string.Equals(Path.GetExtension(name).TrimStart('.'), extension,
                                StringComparison.OrdinalIgnoreCase))
                            found.Add(name);
                    }
                }
            });
        }

        foreach (var entry in entries.Values)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                continue;
            ReadMetadata(entry, limits, provider, operation, cancellationToken, reader =>
            {
                if (reader.LocalName != "Relationship" || reader.NamespaceURI != RelationshipsNamespace
                    || !IsPivotRelationship(reader.GetAttribute("Type")))
                    return;
                if (string.Equals(reader.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"透视关系不能引用外部部件: {entry.FullName}");
                AddTarget(ResolveTarget(entry.FullName, reader.GetAttribute("Target")));
            });
        }
        return Array.AsReadOnly(found.OrderBy(name => name, StringComparer.Ordinal).ToArray());

        void AddTarget(string name)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(name) || !entries.ContainsKey(name))
                throw new InvalidDataException($"透视模板缺少引用部件: {name}");
            found.Add(name);
        }
    }

    /// <summary>
    /// 判断内容类型是否属于透视部件。
    /// </summary>
    /// <param name="contentType">部件内容类型。</param>
    /// <returns>是否属于透视部件。</returns>
    private static bool IsPivotContentType(string contentType) =>
        contentType == ContentTypePrefix + "pivotTable+xml"
        || contentType == ContentTypePrefix + "pivotCacheDefinition+xml"
        || contentType == ContentTypePrefix + "pivotCacheRecords+xml";

    /// <summary>
    /// 判断关系类型是否指向透视部件。
    /// </summary>
    /// <param name="type">关系类型。</param>
    /// <returns>是否属于透视关系。</returns>
    private static bool IsPivotRelationship(string type)
    {
        if (type == null)
            return false;
        var prefix = type.StartsWith(TransitionalRelationshipPrefix, StringComparison.Ordinal)
            ? TransitionalRelationshipPrefix
            : type.StartsWith(StrictRelationshipPrefix, StringComparison.Ordinal)
                ? StrictRelationshipPrefix : null;
        return prefix != null && type[prefix.Length..] is "pivotTable" or "pivotCacheDefinition" or "pivotCacheRecords";
    }

    /// <summary>
    /// 解析关系拥有者的相对目标路径。
    /// </summary>
    /// <param name="relationshipPart">关系部件名称。</param>
    /// <param name="target">关系声明的目标。</param>
    /// <returns>实际包内部件名称。</returns>
    internal static string ResolveTarget(string relationshipPart, string target)
    {
        var name = relationshipPart.Replace('\\', '/').TrimStart('/');
        var separator = name.LastIndexOf("/_rels/", StringComparison.Ordinal);
        var owner = name == "_rels/.rels" ? string.Empty
            : separator >= 0 ? name[..(separator + 1)] + name[(separator + 7)..^5]
            : name.StartsWith("_rels/", StringComparison.Ordinal) && name.EndsWith(".rels", StringComparison.Ordinal)
                ? name[6..^5]
            : throw new InvalidDataException($"透视关系部件路径无效: {relationshipPart}");
        var origin = new Uri("http://bing-offices.local/" + owner, UriKind.Absolute);
        if (string.IsNullOrWhiteSpace(target) || target.StartsWith("//", StringComparison.Ordinal)
            || (!target.StartsWith("/", StringComparison.Ordinal)
                && Uri.TryCreate(target, UriKind.Absolute, out _))
            || !Uri.TryCreate(origin, target, out var resolved)
            || resolved.Scheme != origin.Scheme || resolved.Host != origin.Host
            || !string.IsNullOrEmpty(resolved.Query) || !string.IsNullOrEmpty(resolved.Fragment))
            throw new InvalidDataException($"透视关系目标无效: {target}");
        return resolved.GetComponents(UriComponents.Path, UriFormat.UriEscaped);
    }

    /// <summary>
    /// 扫描包元数据中的元素。
    /// </summary>
    /// <param name="entry">元数据部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="visit">元素处理委托。</param>
    internal static void ReadMetadata(ZipArchiveEntry entry, CancellationToken cancellationToken,
        Action<XmlReader> visit)
        => ReadMetadata(entry, new ExcelResourceLimits
        {
            MaxXmlCharacters = null,
            MaxXmlDepth = null
        }, "Core", BingOfficesOperation.Import, cancellationToken, visit);

    /// <summary>
    /// 扫描包元数据中的元素并应用 XML 资源限制。
    /// </summary>
    /// <param name="entry">元数据部件。</param>
    /// <param name="limits">XML 字符数和嵌套深度限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="visit">元素处理委托。</param>
    /// <exception cref="BingOfficesResourceLimitException">XML 字符数或嵌套深度超过限制。</exception>
    internal static void ReadMetadata(ZipArchiveEntry entry, ExcelResourceLimits limits,
        string provider, BingOfficesOperation operation, CancellationToken cancellationToken,
        Action<XmlReader> visit)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersFromEntities = 0,
                MaxCharactersInDocument = 0
            });
            long characters = 0;
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (limits.MaxXmlDepth.HasValue && reader.Depth > limits.MaxXmlDepth.Value)
                    throw ResourceLimit($"XLSX XML 嵌套深度超过限制: {entry.FullName}", provider, operation);
                var nodeCharacters = (long)(reader.Name?.Length ?? 0) + (reader.Value?.Length ?? 0);
                if (reader.HasAttributes)
                {
                    for (var index = 0; index < reader.AttributeCount; index++)
                    {
                        reader.MoveToAttribute(index);
                        nodeCharacters += (reader.Name?.Length ?? 0) + (reader.Value?.Length ?? 0);
                    }
                    reader.MoveToElement();
                }
                if (limits.MaxXmlCharacters.HasValue
                    && nodeCharacters > limits.MaxXmlCharacters.Value - characters)
                    throw ResourceLimit($"XLSX XML 字符数量超过限制: {entry.FullName}", provider, operation);
                characters += nodeCharacters;
                if (reader.NodeType == XmlNodeType.Element)
                    visit(reader);
            }
        }
        catch (XmlException exception)
        {
            throw new InvalidDataException($"透视模板元数据无效: {entry.FullName}", exception);
        }
    }

    /// <summary>
    /// 创建资源限制异常。
    /// </summary>
    /// <param name="message">资源限制说明。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <returns>带有预检阶段上下文的资源限制异常。</returns>
    private static BingOfficesResourceLimitException ResourceLimit(string message, string provider,
        BingOfficesOperation operation) => new(message, provider: provider,
            operation: operation, stage: BingOfficesStage.Preflight);
}
