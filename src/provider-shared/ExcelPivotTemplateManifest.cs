using System.Globalization;
using System.IO.Compression;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;

namespace Bing.Offices.Internals;

/// <summary>
/// 已有透视模板的不可变依赖清单。
/// </summary>
/// <remarks>只读取关系和结构元数据；缓存记录及工作表数据由后续执行阶段处理。</remarks>
internal sealed class ExcelPivotTemplateManifest
{
    /// <summary>
    /// ClosedXML 生成的透视表扩展标识。
    /// </summary>
    private const string ClosedXmlPivotExtensionUri =
        "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}";

    /// <summary>
    /// ClosedXML 透视表扩展命名空间。
    /// </summary>
    private const string ClosedXmlPivotExtensionNamespace =
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    /// <summary>
    /// DrawingML 图表命名空间。
    /// </summary>
    private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>
    /// Office 图表扩展命名空间。
    /// </summary>
    private const string ChartExtensionNamespace = "http://schemas.microsoft.com/office/drawing/2012/chart";

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateManifest" /> 类型的实例。
    /// </summary>
    /// <param name="workbookPart">工作簿部件。</param>
    /// <param name="sheetNames">模板工作表名称。</param>
    /// <param name="parts">发现的透视部件。</param>
    /// <param name="entries">透视表快照。</param>
    /// <param name="caches">缓存快照。</param>
    /// <param name="relationships">关系快照。</param>
    private ExcelPivotTemplateManifest(string workbookPart, IEnumerable<string> sheetNames, IEnumerable<string> parts,
        IEnumerable<ExcelPivotTemplateEntry> entries, IEnumerable<ExcelPivotTemplateCache> caches,
        IReadOnlyDictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>> relationships,
        IEnumerable<string> relatedFeatures)
    {
        WorkbookPart = workbookPart;
        SheetNames = Array.AsReadOnly(sheetNames.ToArray());
        Parts = Array.AsReadOnly(parts.ToArray());
        Entries = Array.AsReadOnly(entries.ToArray());
        Caches = Array.AsReadOnly(caches.ToArray());
        Relationships = relationships;
        RelatedFeatures = Array.AsReadOnly(relatedFeatures.ToArray());
    }

    /// <summary>
    /// 获取工作簿部件。
    /// </summary>
    internal string WorkbookPart { get; }

    /// <summary>
    /// 获取模板中的工作表名称。
    /// </summary>
    internal IReadOnlyList<string> SheetNames { get; }

    /// <summary>
    /// 获取发现的透视部件。
    /// </summary>
    internal IReadOnlyList<string> Parts { get; }

    /// <summary>
    /// 获取透视表描述。
    /// </summary>
    internal IReadOnlyList<ExcelPivotTemplateEntry> Entries { get; }

    /// <summary>
    /// 获取缓存描述。
    /// </summary>
    internal IReadOnlyList<ExcelPivotTemplateCache> Caches { get; }

    /// <summary>
    /// 获取按拥有者索引的原始关系。
    /// </summary>
    internal IReadOnlyDictionary<string, IReadOnlyList<ExcelPivotTemplateRelationship>> Relationships { get; }

    /// <summary>
    /// 获取与模板透视相关且当前不支持更新的功能。
    /// </summary>
    internal IReadOnlyList<string> RelatedFeatures { get; }

    /// <summary>
    /// 读取已有透视模板的依赖清单。
    /// </summary>
    /// <param name="source">调用方持有的可读取、可定位模板流。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>保留原始部件和关系标识的只读清单。</returns>
    /// <remarks>读取成功、失败或取消时均恢复输入位置，并保持流打开。</remarks>
    internal static ExcelPivotTemplateManifest Read(Stream source, CancellationToken token = default)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek)
            throw new ArgumentException("透视模板清单要求可读取且可定位的流。", nameof(source));
        var position = source.Position;
        try
        {
            token.ThrowIfCancellationRequested();
            source.Position = 0;
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            return Read(archive, token);
        }
        finally
        {
            source.Position = position;
        }
    }

    /// <summary>
    /// 读取已有透视模板的依赖清单。
    /// </summary>
    /// <param name="source">调用方持有的模板流。</param>
    /// <param name="limits">ZIP 和 XML 资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>通过资源预检的只读清单。</returns>
    /// <remarks>允许工作簿位于非标准目录；实际根部件由关系清单解析。</remarks>
    internal static ExcelPivotTemplateManifest Read(Stream source, ExcelResourceLimits limits,
        string provider, CancellationToken token = default)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead || !source.CanSeek)
            throw new ArgumentException("透视模板清单要求可读取且可定位的流。", nameof(source));
        if (limits == null)
            throw new ArgumentNullException(nameof(limits));
        limits.Validate();
        var position = source.Position;
        try
        {
            token.ThrowIfCancellationRequested();
            if (limits.MaxInputBytes.HasValue && source.Length > limits.MaxInputBytes.Value)
                throw new BingOfficesResourceLimitException("透视模板输入大小超过限制。", provider: provider);
            ExcelXlsxZipPreflight.Validate(source, limits, provider, requireZip: true,
                cancellationToken: token, enforceSheetLimit: false, enforceCellLimits: false,
                requireWorkbookXml: false);
            source.Position = 0;
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            return Read(archive, limits, provider, BingOfficesOperation.Import, token);
        }
        finally
        {
            source.Position = position;
        }
    }

    /// <summary>
    /// 读取已有透视模板的依赖清单。
    /// </summary>
    /// <param name="archive">调用方持有的模板包。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>保留原始部件和关系标识的只读清单。</returns>
    internal static ExcelPivotTemplateManifest Read(ZipArchive archive, CancellationToken token = default)
        => ReadCore(archive, null, "Core", BingOfficesOperation.Import, token);

    /// <summary>
    /// 读取透视模板依赖清单并限制元数据解析资源。
    /// </summary>
    /// <param name="archive">调用方持有的模板包。</param>
    /// <param name="limits">ZIP 和 XML 资源限制。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="operation">触发预检的操作类型。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>通过资源预检的只读清单。</returns>
    internal static ExcelPivotTemplateManifest Read(ZipArchive archive, ExcelResourceLimits limits,
        string provider, BingOfficesOperation operation, CancellationToken token = default)
    {
        if (limits == null)
            throw new ArgumentNullException(nameof(limits));
        limits.Validate();
        return ReadCore(archive, limits, provider, operation, token);
    }

    /// <summary>
    /// 使用指定的 XML 限制读取透视模板依赖清单。
    /// </summary>
    private static ExcelPivotTemplateManifest ReadCore(ZipArchive archive, ExcelResourceLimits limits,
        string provider, BingOfficesOperation operation, CancellationToken token)
    {
        if (archive == null)
            throw new ArgumentNullException(nameof(archive));
        var discovered = limits == null
            ? ExcelPivotTemplateParts.Find(archive, token)
            : ExcelPivotTemplateParts.Find(archive, limits, provider, operation, token);
        var package = limits == null
            ? new ExcelPivotTemplatePackage(archive, token)
            : new ExcelPivotTemplatePackage(archive, token, limits, provider, operation);
        var roots = package.Relationships(string.Empty, token)
            .Where(item => ExcelPivotTemplatePackage.IsRelationship(item.Type, "officeDocument")).ToArray();
        if (roots.Length != 1)
            throw new InvalidDataException("模板必须包含唯一的工作簿根关系。");
        var workbook = package.RequireRelationship(string.Empty, roots[0].Id, "officeDocument", token).PartName;
        var sheets = new List<(string Name, string RelationshipId)>();
        var cacheReferences = new List<(uint Id, string RelationshipId)>();
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var cacheIds = new HashSet<uint>();
        var workbookRoot = false;
        package.Read(workbook, token, reader =>
        {
            if (!IsSpreadsheet(reader))
                return;
            if (reader.Depth == 0)
                workbookRoot = reader.LocalName == "workbook";
            if (reader.LocalName == "sheet" && reader.Depth == 2)
            {
                var name = ExcelPivotTemplatePackage.Required(reader.GetAttribute("name"), workbook);
                if (!sheetNames.Add(name))
                    throw new InvalidDataException($"模板包含重复工作表名称: {name}");
                sheets.Add((name, RelationshipId(reader, workbook)));
            }
            if (reader.LocalName == "pivotCache" && reader.Depth == 2)
            {
                var id = CacheId(reader.GetAttribute("cacheId"), workbook);
                if (!cacheIds.Add(id))
                    throw new InvalidDataException($"模板缓存标识重复: {id}");
                cacheReferences.Add((id, RelationshipId(reader, workbook)));
            }
        });
        if (!workbookRoot)
            throw new InvalidDataException($"模板工作簿根元素无效: {workbook}");

        var pivots = new List<(string Name, string SheetName, string SheetPart, string RelationshipId,
            string Part, uint CacheId, string Target, bool HasUnverifiedDependencies)>();
        var pivotNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pivotParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in sheets)
        {
            token.ThrowIfCancellationRequested();
            var sheetPart = package.RequireRelationship(workbook, sheet.RelationshipId, "worksheet", token).PartName;
            var declaredReferences = new HashSet<string>(StringComparer.Ordinal);
            var worksheetRoot = false;
            package.Read(sheetPart, token, reader =>
            {
                if (IsSpreadsheet(reader) && reader.Depth == 0)
                    worksheetRoot = reader.LocalName == "worksheet";
                if (IsSpreadsheet(reader) && reader.LocalName == "pivotTablePart" && reader.Depth == 2)
                    if (!declaredReferences.Add(RelationshipId(reader, sheetPart)))
                        throw new InvalidDataException($"工作表透视标识重复: {sheetPart}");
            });
            if (!worksheetRoot)
                throw new InvalidDataException($"工作表根元素无效: {sheetPart}");
            foreach (var id in declaredReferences)
                package.RequireRelationship(sheetPart, id, "pivotTable", token);
            // NPOI 的实际工作表仅写包关系；归属以 OPC 关系为准，正文引用存在时仍须有效。
            var references = package.Relationships(sheetPart, token)
                .Where(item => ExcelPivotTemplatePackage.IsRelationship(item.Type, "pivotTable"))
                .Select(item => item.Id);
            foreach (var id in references)
            {
                var part = package.RequireRelationship(sheetPart, id, "pivotTable", token).PartName;
                if (!pivotParts.Add(part))
                    throw new InvalidDataException($"透视部件被重复引用: {part}");
                string name = null;
                string target = null;
                uint? cacheId = null;
                package.Read(part, token, reader =>
                {
                    if (!IsSpreadsheet(reader))
                        return;
                    if (reader.Depth == 0 && reader.LocalName == "pivotTableDefinition")
                    {
                        name = ExcelPivotTemplatePackage.Required(reader.GetAttribute("name"), part);
                        cacheId = CacheId(reader.GetAttribute("cacheId"), part);
                    }
                    else if (reader.Depth == 1 && reader.LocalName == "location")
                        target = ExcelPivotTemplatePackage.Required(reader.GetAttribute("ref"), part);
                });
                if (name == null || cacheId == null || target == null || !cacheIds.Contains(cacheId.Value))
                    throw new InvalidDataException($"透视定义或缓存标识无效: {part}");
                if (!pivotNames.Add(name))
                    throw new InvalidDataException($"模板包含重复透视名称: {name}");
                pivots.Add((name, sheet.Name, sheetPart, id, part, cacheId.Value, target,
                    HasUnverifiedMarkup(package, part, token)
                    || HasUnexpectedRelationship(package, part, "pivotCacheDefinition", token)));
            }
        }

        var caches = new Dictionary<uint, ExcelPivotTemplateCache>();
        var referencedParts = new HashSet<string>(pivotParts, StringComparer.OrdinalIgnoreCase);
        foreach (var reference in cacheReferences)
        {
            token.ThrowIfCancellationRequested();
            var definition = package.RequireRelationship(workbook, reference.RelationshipId,
                "pivotCacheDefinition", token).PartName;
            referencedParts.Add(definition);
            string recordId = null;
            string sourceType = null;
            string sourceSheet = null;
            string sourceRef = null;
            string sourceName = null;
            var fields = new List<string>();
            var cacheRoot = false;
            package.Read(definition, token, reader =>
            {
                if (!IsSpreadsheet(reader))
                    return;
                if (reader.Depth == 0 && reader.LocalName == "pivotCacheDefinition")
                {
                    cacheRoot = true;
                    recordId = OptionalRelationshipId(reader);
                }
                else if (reader.Depth == 1 && reader.LocalName == "cacheSource")
                    sourceType = ExcelPivotTemplatePackage.Required(reader.GetAttribute("type"), definition);
                else if (reader.Depth == 2 && reader.LocalName == "worksheetSource")
                {
                    sourceSheet = reader.GetAttribute("sheet");
                    sourceRef = reader.GetAttribute("ref");
                    sourceName = reader.GetAttribute("name");
                }
                else if (reader.Depth == 2 && reader.LocalName == "cacheField")
                    fields.Add(ExcelPivotTemplatePackage.Required(reader.GetAttribute("name"), definition));
            });
            if (!cacheRoot || sourceType == null)
                throw new InvalidDataException($"缓存定义无效: {definition}");
            if (sourceType == "worksheet" && sourceName == null)
                ExcelPivotTemplateSource.Validate(sourceSheet, sourceRef, sheetNames, definition);
            var hasUnverifiedDependencies = !string.IsNullOrWhiteSpace(sourceName)
                || HasUnverifiedMarkup(package, definition, token)
                || HasUnexpectedRelationship(package, definition, "pivotCacheRecords", token);
            string records = null;
            if (recordId != null)
            {
                records = package.RequireRelationship(definition, recordId, "pivotCacheRecords", token).PartName;
                referencedParts.Add(records);
                hasUnverifiedDependencies |= HasUnverifiedMarkup(package, records, token)
                    || HasUnexpectedRelationship(package, records, null, token);
            }
            foreach (var pivot in pivots.Where(item => item.CacheId == reference.Id))
            {
                var links = package.Relationships(pivot.Part, token)
                    .Where(item => ExcelPivotTemplatePackage.IsRelationship(item.Type, "pivotCacheDefinition")).ToArray();
                if (links.Length != 1)
                    throw new InvalidDataException($"透视表必须包含唯一缓存关系: {pivot.Part}");
                var linkedDefinition = package.RequireRelationship(pivot.Part, links[0].Id,
                    "pivotCacheDefinition", token).PartName;
                if (!string.Equals(linkedDefinition, definition, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException($"透视表缓存关系与 cacheId 不一致: {pivot.Part}");
            }
            caches.Add(reference.Id, new ExcelPivotTemplateCache(reference.Id, reference.RelationshipId,
                definition, records, recordId, sourceType, sourceSheet, sourceRef, sourceName, fields,
                pivots.Where(item => item.CacheId == reference.Id).Select(item => item.Part),
                hasUnverifiedDependencies));
        }
        foreach (var part in discovered)
            if (!part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) && !referencedParts.Contains(part))
                throw new InvalidDataException($"透视模板存在未关联部件: {part}");
        var entries = pivots.Select(item => new ExcelPivotTemplateEntry(item.Name, item.SheetName,
            item.SheetPart, item.RelationshipId, item.Part, item.Target, caches[item.CacheId],
            item.HasUnverifiedDependencies));
        var relatedFeatures = FindRelatedFeatures(package, workbook, sheets, token);
        return new ExcelPivotTemplateManifest(workbook, sheets.Select(item => item.Name), discovered, entries,
            caches.Values, package.SnapshotRelationships(), relatedFeatures);
    }

    /// <summary>
    /// 判断元素是否属于工作簿 XML。
    /// </summary>
    /// <param name="reader">当前元素读取器。</param>
    /// <returns>是否属于过渡或严格格式。</returns>
    private static bool IsSpreadsheet(XmlReader reader) =>
        reader.NamespaceURI == "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        || reader.NamespaceURI == "http://purl.oclc.org/ooxml/spreadsheetml/main";

    /// <summary>
    /// 获取必填关系标识。
    /// </summary>
    /// <param name="reader">当前元素读取器。</param>
    /// <param name="part">所属部件。</param>
    /// <returns>关系标识。</returns>
    private static string RelationshipId(XmlReader reader, string part) =>
        ExcelPivotTemplatePackage.Required(OptionalRelationshipId(reader), part);

    /// <summary>
    /// 获取可选关系标识。
    /// </summary>
    /// <param name="reader">当前元素读取器。</param>
    /// <returns>过渡或严格格式的关系标识。</returns>
    private static string OptionalRelationshipId(XmlReader reader) =>
        reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        ?? reader.GetAttribute("id", "http://purl.oclc.org/ooxml/officeDocument/relationships");

    /// <summary>
    /// 判断部件是否包含当前更新路径无法安全重建的扩展标记。
    /// </summary>
    /// <param name="package">模板部件包。</param>
    /// <param name="part">待检查的部件。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>包含扩展标记时返回 <see langword="true" />。</returns>
    private static bool HasUnverifiedMarkup(ExcelPivotTemplatePackage package, string part,
        CancellationToken token)
    {
        var found = false;
        var extensionDepth = -1;
        var knownExtension = false;
        package.Read(part, token, reader =>
        {
            if (reader.LocalName == "ext")
            {
                extensionDepth = reader.Depth;
                knownExtension = string.Equals(reader.GetAttribute("uri"),
                    ClosedXmlPivotExtensionUri, StringComparison.OrdinalIgnoreCase);
                if (!knownExtension)
                    found = true;
                return;
            }

            if (extensionDepth >= 0 && reader.Depth <= extensionDepth)
                extensionDepth = -1;
            if (extensionDepth < 0)
                return;
            if (!knownExtension || reader.Depth != extensionDepth + 1
                || reader.LocalName != "pivotTableDefinition"
                || reader.NamespaceURI != ClosedXmlPivotExtensionNamespace)
                found = true;
        });
        return found;
    }

    /// <summary>
    /// 判断部件是否包含未纳入当前更新契约的关系。
    /// </summary>
    /// <param name="package">模板部件包。</param>
    /// <param name="owner">关系拥有者。</param>
    /// <param name="expectedKind">允许的关系类型；记录部件不允许关系时传入 <see langword="null" />。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>存在其它关系时返回 <see langword="true" />。</returns>
    private static bool HasUnexpectedRelationship(ExcelPivotTemplatePackage package, string owner,
        string expectedKind, CancellationToken token) =>
        package.Relationships(owner, token).Any(item => expectedKind == null
            || !ExcelPivotTemplatePackage.IsRelationship(item.Type, expectedKind));

    /// <summary>
    /// 识别透视模板中的关联图表、切片器和时间线。
    /// </summary>
    /// <param name="package">模板部件包。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>按名称排序的关联功能。</returns>
    private static IReadOnlyList<string> FindRelatedFeatures(ExcelPivotTemplatePackage package,
        string workbookPart, IReadOnlyList<(string Name, string RelationshipId)> sheets, CancellationToken token)
    {
        var features = new HashSet<string>(StringComparer.Ordinal);
        foreach (var relationship in package.Relationships(workbookPart, token))
        {
            token.ThrowIfCancellationRequested();
            var kind = RelationshipKind(relationship.Type);
            if (string.Equals(kind, "slicerCache", StringComparison.OrdinalIgnoreCase)
                && (relationship.External
                    || HasPivotSlicerReferences(package, relationship.PartName, token)))
            {
                features.Add("切片器");
            }
            else if (string.Equals(kind, "timelineCache", StringComparison.OrdinalIgnoreCase))
            {
                if (!relationship.External)
                    package.Read(relationship.PartName, token, _ => { });
                features.Add("时间线");
            }
        }

        foreach (var sheet in sheets)
        {
            token.ThrowIfCancellationRequested();
            var sheetPart = package.RequireRelationship(workbookPart, sheet.RelationshipId, "worksheet", token).PartName;
            foreach (var drawing in package.Relationships(sheetPart, token)
                         .Where(item => string.Equals(RelationshipKind(item.Type), "drawing",
                             StringComparison.OrdinalIgnoreCase) && !item.External))
            {
                token.ThrowIfCancellationRequested();
                package.Read(drawing.PartName, token, _ => { });
                foreach (var chart in package.Relationships(drawing.PartName, token)
                             .Where(item => string.Equals(RelationshipKind(item.Type), "chart",
                                 StringComparison.OrdinalIgnoreCase) && !item.External))
                    if (IsPivotChart(package, chart.PartName, token))
                        features.Add("透视图");
            }
        }

        return Array.AsReadOnly(features.OrderBy(item => item, StringComparer.Ordinal).ToArray());
    }

    /// <summary>
    /// 获取关系类型的末段名称。
    /// </summary>
    /// <param name="type">关系类型。</param>
    /// <returns>关系类型名称。</returns>
    private static string RelationshipKind(string type)
    {
        var separator = type.LastIndexOf('/');
        return separator >= 0 ? type[(separator + 1)..] : type;
    }

    /// <summary>
    /// 判断图表部件是否声明了透视数据源。
    /// </summary>
    /// <param name="package">模板部件包。</param>
    /// <param name="part">图表部件。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>图表包含透视数据源时返回 <see langword="true" />。</returns>
    private static bool IsPivotChart(ExcelPivotTemplatePackage package, string part,
        CancellationToken token)
    {
        var isChart = false;
        var hasPivotSource = false;
        package.Read(part, token, reader =>
        {
            if (reader.Depth == 0 && reader.LocalName == "chartSpace"
                && reader.NamespaceURI == ChartNamespace)
                isChart = true;
            if (reader.LocalName == "pivotSource"
                && reader.NamespaceURI is ChartNamespace or ChartExtensionNamespace)
                hasPivotSource = true;
        });
        return isChart && hasPivotSource;
    }

    /// <summary>
    /// 判断切片器缓存是否关联透视表。
    /// </summary>
    /// <param name="package">模板部件包。</param>
    /// <param name="part">切片器缓存部件。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>缓存声明透视表引用时返回 <see langword="true" />。</returns>
    private static bool HasPivotSlicerReferences(ExcelPivotTemplatePackage package, string part,
        CancellationToken token)
    {
        var hasPivotTableReferences = false;
        package.Read(part, token, reader =>
        {
            if (reader.LocalName == "pivotTables" || reader.LocalName == "pivotTable")
                hasPivotTableReferences = true;
        });
        return hasPivotTableReferences;
    }

    /// <summary>
    /// 解析缓存标识。
    /// </summary>
    /// <param name="value">缓存标识文本。</param>
    /// <param name="part">所属部件。</param>
    /// <returns>有效缓存标识。</returns>
    private static uint CacheId(string value, string part) =>
        uint.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var id)
            ? id : throw new InvalidDataException($"透视模板缓存标识无效: {part}/{value}");
}
