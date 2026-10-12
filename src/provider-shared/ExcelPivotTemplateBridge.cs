using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;

namespace Bing.Offices.Internals;

/// <summary>
/// 在 Provider 工作簿写入前后桥接已有透视部件。
/// </summary>
/// <remarks>
/// 中间模板移除透视关系以避免 DOM Provider 改写未知部件；业务 Sheet 写入完成后，
/// 保留模式回挂原始透视部件，更新模式回挂 Provider 生成的缓存和透视关系。该桥接不执行客户端刷新。
/// </remarks>
internal sealed class ExcelPivotTemplateBridge : IDisposable
{
    private const long MaximumXmlCharacters = 64L * 1024 * 1024;
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeDocumentRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// <summary>
    /// 更新时从模板复制的透视布局属性。
    /// </summary>
    private static readonly string[] PivotLayoutAttributes =
    {
        "dataCaption", "grandTotalCaption", "rowHeaderCaption", "colHeaderCaption",
        "fieldListSortAscending", "autoFormatId", "useAutoFormatting", "applyNumberFormats",
        "applyBorderFormats", "applyFontFormats", "applyPatternFormats", "applyAlignmentFormats",
        "applyWidthHeightFormats", "showError", "showMissing", "pageWrap", "pageOverThenDown",
        "subtotalHiddenItems", "rowGrandTotals", "colGrandTotals", "fieldPrintTitles",
        "itemPrintTitles", "mergeItem", "showDropZones", "showDrill", "printDrill",
        "showMemberPropertyTips", "showDataTips", "showDataDropDown", "showItems",
        "showMultipleLabel", "showEmptyRow",
        "showEmptyCol", "showHeaders", "compact", "compactData", "outline", "outlineData",
        "preserveFormatting"
    };
    /// <summary>
    /// 更新时从模板复制的字段索引格式元素。
    /// </summary>
    private static readonly string[] PivotLayoutElements =
    {
        "formats", "conditionalFormats", "chartFormats"
    };

    private readonly ExcelPivotTemplateSnapshot _snapshot;
    private readonly FileStream _strippedTemplate;
    private readonly string _provider;

    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplateBridge" /> 类型的实例。
    /// </summary>
    /// <param name="snapshot">原始模板快照。</param>
    /// <param name="strippedTemplate">已移除透视关系的中间模板。</param>
    /// <param name="provider">提供程序名称。</param>
    private ExcelPivotTemplateBridge(ExcelPivotTemplateSnapshot snapshot,
        FileStream strippedTemplate, string provider)
    {
        _snapshot = snapshot;
        _strippedTemplate = strippedTemplate;
        _provider = provider;
    }

    /// <summary>
    /// 获取供 Provider 加载的中间模板流。
    /// </summary>
    internal Stream Template => _strippedTemplate;

    /// <summary>
    /// 获取是否包含已有透视更新。
    /// </summary>
    internal bool HasUpdates => _snapshot.Plan.Updates.Count != 0;

    /// <summary>
    /// 创建供 Provider 生成更新缓存的临时透视定义。
    /// </summary>
    /// <returns>目标指向临时工作表的透视定义。</returns>
    internal IReadOnlyList<ExcelPivotTableDefinition> CreateProviderUpdateDefinitions()
    {
        if (!HasUpdates)
            return Array.Empty<ExcelPivotTableDefinition>();
        ValidateSupportedUpdateShape(_snapshot.Plan, _provider);
        var names = new HashSet<string>(_snapshot.Plan.Manifest.SheetNames,
            StringComparer.OrdinalIgnoreCase);
        var definitions = new List<ExcelPivotTableDefinition>(_snapshot.Plan.Updates.Count);
        for (var index = 0; index < _snapshot.Plan.Updates.Count; index++)
        {
            var update = _snapshot.Plan.Updates[index].Definition;
            var target = TemporarySheetName(index + 1);
            if (!names.Add(target))
                throw UnsupportedUpdate("模板已占用透视更新临时工作表名称。", BingOfficesStage.Preflight);
            definitions.Add(new ExcelPivotTableDefinition
            {
                Name = update.Name,
                SourceSheet = update.SourceSheet,
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = update.SourceRange.StartRow,
                    StartColumn = update.SourceRange.StartColumn,
                    EndRow = update.SourceRange.EndRow,
                    EndColumn = update.SourceRange.EndColumn
                },
                TargetSheet = target,
                TargetRow = update.TargetRow,
                TargetColumn = update.TargetColumn,
                RowField = update.RowField,
                ValueField = update.ValueField,
                RowFields = update.RowFields,
                ColumnFields = update.ColumnFields,
                ValueFields = update.ValueFields,
                Filters = update.Filters,
                MaxExpandedCells = update.MaxExpandedCells
            });
        }
        return definitions.AsReadOnly();
    }

    /// <summary>
    /// 校验当前阶段支持的透视更新形状。
    /// </summary>
    /// <param name="plan">模板执行计划。</param>
    /// <param name="provider">提供程序名称。</param>
    private static void ValidateSupportedUpdateShape(ExcelPivotTemplatePlan plan, string provider)
    {
        if (plan.Updates.Count == 0)
            return;
        var names = new HashSet<string>(plan.Manifest.SheetNames,
            StringComparer.OrdinalIgnoreCase);
        for (var index = 1; index <= plan.Updates.Count; index++)
            if (names.Contains(TemporarySheetName(index)))
                throw new BingOfficesUnsupportedFeatureException(
                    $"{provider}: 模板已占用透视更新临时工作表名称。", provider: provider,
                    operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 获取透视更新的临时工作表名称。
    /// </summary>
    /// <param name="index">从一开始的更新序号。</param>
    /// <returns>临时工作表名称。</returns>
    private static string TemporarySheetName(int index) =>
        "__BingPivotUpdate_" + index.ToString("000", System.Globalization.CultureInfo.InvariantCulture);

    /// <summary>
    /// 判断请求是否需要透视部件桥接。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <returns>带业务写入且需要透视部件桥接的请求。</returns>
    internal static bool IsRequired(ExcelWorkbookExportRequest request) =>
        ExcelPivotTemplatePreserver.IsPreserveTemplate(request)
        && !ExcelPivotTemplatePreserver.IsPreserveOnly(request);

    /// <summary>
    /// 创建透视部件桥接。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>已完成模板预检的桥接；无需桥接时返回 <see langword="null" />。</returns>
    internal static ExcelPivotTemplateBridge Create(ExcelWorkbookExportRequest request,
        string provider, CancellationToken token)
    {
        if (!IsRequired(request))
            return null;
        var snapshot = ExcelPivotTemplateSnapshot.Create(request,
            request.ResourceLimits ?? new ExcelResourceLimits(), provider, token);
        try
        {
            ValidateSupportedUpdateShape(snapshot.Plan, provider);
            var stripped = Strip(snapshot.Source, snapshot.Plan.Manifest, token,
                snapshot.Limits.MaxXmlCharacters);
            return new ExcelPivotTemplateBridge(snapshot, stripped, provider);
        }
        catch
        {
            snapshot.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 将 Provider 生成的工作簿与原始透视部件合并。
    /// </summary>
    /// <param name="generated">Provider 写入的中间输出。</param>
    /// <param name="destination">最终目标流。</param>
    /// <param name="token">取消令牌。</param>
    internal void Merge(Stream generated, Stream destination, CancellationToken token)
    {
        try
        {
            token.ThrowIfCancellationRequested();
            ValidateGeneratedPackage(generated, token);
            using var output = ReadArchive(generated, token);
            if (HasUpdates)
            {
                if (_snapshot.Plan.PreservedEntries.Count == 0)
                    MergeUpdates(output, generated, token);
                else
                    MergePartialUpdates(output, generated, token);
                using var updated = CreateArchive(output, token);
                updated.Position = 0;
                Copy(updated, destination, token);
                return;
            }
            _snapshot.Source.Position = 0;
            using var original = ReadArchive(_snapshot.Source, token);
            var outputPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
            var originalPackage = ReadPackage(original, token, _snapshot.Limits.MaxXmlCharacters);
            var manifest = _snapshot.Plan.Manifest;

            if (manifest.Parts.Count > 0)
            {
                var generatedManifest = ReadGeneratedManifest(output, token);
                var preservedParts = new HashSet<string>(
                    manifest.Parts.Select(Normalize), StringComparer.OrdinalIgnoreCase);
                foreach (var part in manifest.Parts)
                    preservedParts.Add(RelationshipPart(part));
                foreach (var owner in manifest.Relationships.Keys)
                    if (manifest.Parts.Contains(owner, StringComparer.OrdinalIgnoreCase))
                        preservedParts.Add(RelationshipPart(owner));

                RenameGeneratedParts(output, generatedManifest, preservedParts, token,
                    _snapshot.Limits.MaxXmlCharacters);
                generatedManifest = ReadGeneratedManifest(output, token);
                var cacheIds = CreateCacheIdMap(manifest.Caches, generatedManifest.Caches);
                RemapGeneratedCacheIds(output, generatedManifest, cacheIds, token,
                    _snapshot.Limits.MaxXmlCharacters);
                outputPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
            }

            CopyPivotParts(original, output, manifest);
            MergeContentTypes(original, output, manifest, token,
                _snapshot.Limits.MaxXmlCharacters);
            MergeWorkbook(original, output, originalPackage, outputPackage, manifest, token);
            MergeWorksheets(output, originalPackage, outputPackage, manifest, token);

            using var merged = CreateArchive(output, token, CompressionLevel.Optimal);
            Copy(merged, destination, token);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesExportException("透视模板部件合并失败。", exception, _provider,
                BingOfficesStage.Write);
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        _strippedTemplate.Dispose();
        _snapshot.Dispose();
    }

    /// <summary>
    /// 从模板中移除透视部件及其工作簿、工作表关系。
    /// </summary>
    /// <param name="source">原始模板流。</param>
    /// <param name="manifest">模板透视清单。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="maxXmlCharacters">XML 部件允许的最大字符数。</param>
    /// <returns>可供 Provider 加载的中间模板。</returns>
    private static FileStream Strip(Stream source, ExcelPivotTemplateManifest manifest,
        CancellationToken token, long? maxXmlCharacters)
    {
        var originalPosition = source.Position;
        try
        {
            source.Position = 0;
            using var input = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var pivotParts = new HashSet<string>(manifest.Parts, StringComparer.OrdinalIgnoreCase);
            var sheetParts = new HashSet<string>(manifest.Entries.Select(item => item.SheetPart),
                StringComparer.OrdinalIgnoreCase);
            var output = OpenStrippedTemplateFile();
            try
            {
                using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
                {
                    foreach (var entry in input.Entries)
                    {
                        token.ThrowIfCancellationRequested();
                        var name = Normalize(entry.FullName);
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal)
                            || pivotParts.Contains(name) || IsPivotRelationshipPart(name, pivotParts))
                            continue;
                        byte[] bytes;
                        if (string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                            bytes = EditXml(entry, document => RemoveContentTypes(document, pivotParts),
                                token, maxXmlCharacters);
                        else if (string.Equals(name, manifest.WorkbookPart, StringComparison.OrdinalIgnoreCase))
                            bytes = EditXml(entry, RemoveWorkbookPivotCaches, token, maxXmlCharacters);
                        else if (sheetParts.Contains(name))
                            bytes = EditXml(entry, RemoveWorksheetPivotParts, token, maxXmlCharacters);
                        else if (name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                            bytes = EditXml(entry, document => RemovePivotRelationships(document, pivotParts, name),
                                token, maxXmlCharacters);
                        else
                            bytes = ReadEntry(entry, token);
                        AddEntry(archive, name, bytes);
                    }
                }
                output.Position = 0;
                return output;
            }
            catch
            {
                output.Dispose();
                throw;
            }
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 将原始透视定义、缓存和关系文件复制到 Provider 输出包。
    /// </summary>
    private static void CopyPivotParts(ExcelPivotTemplateArchive original,
        ExcelPivotTemplateArchive output, ExcelPivotTemplateManifest manifest)
    {
        var names = new HashSet<string>(manifest.Parts, StringComparer.OrdinalIgnoreCase);
        foreach (var name in manifest.Parts)
            names.Add(RelationshipPart(name));
        foreach (var owner in manifest.Relationships.Keys)
            if (manifest.Parts.Contains(owner, StringComparer.OrdinalIgnoreCase))
                names.Add(RelationshipPart(owner));
        foreach (var name in names)
        {
            if (!original.TryGetValue(name, out var bytes))
                continue;
            if (output.ContainsKey(name))
                throw new InvalidDataException($"Provider 输出已占用透视部件路径: {name}");
            output[name] = bytes;
        }
    }

    /// <summary>
    /// 合并透视部件的内容类型声明。
    /// </summary>
    private static void MergeContentTypes(ExcelPivotTemplateArchive original,
        ExcelPivotTemplateArchive output, ExcelPivotTemplateManifest manifest,
        CancellationToken token, long? maxXmlCharacters)
    {
        if (!original.TryGetValue("[Content_Types].xml", out var originalBytes)
            || !output.TryGetValue("[Content_Types].xml", out var outputBytes))
            throw new InvalidDataException("透视模板缺少内容类型部件。");
        var source = Load(originalBytes, token, maxXmlCharacters);
        var target = Load(outputBytes, token, maxXmlCharacters);
        var pivotParts = new HashSet<string>(manifest.Parts.Select(Normalize), StringComparer.OrdinalIgnoreCase);
        var existing = target.Root.Elements().Where(item => item.Name.LocalName == "Override")
            .Select(item => Normalize(((string)item.Attribute("PartName"))?.TrimStart('/')))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var item in source.Root.Elements().Where(item => item.Name.LocalName == "Override"))
        {
            token.ThrowIfCancellationRequested();
            var part = Normalize(((string)item.Attribute("PartName"))?.TrimStart('/'));
            if (pivotParts.Contains(part) && existing.Add(part))
                target.Root.Add(new XElement(item));
        }
        output["[Content_Types].xml"] = Save(target);
    }

    /// <summary>
    /// 合并工作簿级透视缓存关系。
    /// </summary>
    private static void MergeWorkbook(ExcelPivotTemplateArchive original,
        ExcelPivotTemplateArchive output, PackageInfo source, PackageInfo target,
        ExcelPivotTemplateManifest manifest, CancellationToken token)
    {
        if (manifest.Caches.Count == 0)
            return;
        var originalCaches = FindElement(source.Workbook.Root, "pivotCaches");
        if (originalCaches == null)
            throw new InvalidDataException("透视清单缺少工作簿缓存声明。");
        var targetCaches = FindElement(target.Workbook.Root, "pivotCaches");
        var addTargetCaches = targetCaches == null;
        targetCaches ??= new XElement(target.Workbook.Root.Name.Namespace + "pivotCaches");
        var targetRelations = EnsureRelationships(target.WorkbookRelationships);
        foreach (var cache in manifest.Caches)
        {
            token.ThrowIfCancellationRequested();
            var originalCache = originalCaches.Elements()
                .FirstOrDefault(item => item.Name.LocalName == "pivotCache"
                    && uint.TryParse((string)item.Attribute("cacheId"), out var id) && id == cache.Id);
            if (originalCache == null)
                throw new InvalidDataException($"透视缓存声明缺少 cacheId: {cache.Id}");
            var relation = FindRelationship(source.WorkbookRelationships, cache.RelationshipId,
                source.WorkbookRelationshipsPart);
            if (relation == null)
                throw new InvalidDataException($"透视缓存关系缺失: {cache.RelationshipId}");
            var id = AddRelationship(targetRelations, relation, cache.DefinitionPart, target.WorkbookPart);
            var copy = new XElement(originalCache);
            SetRelationshipId(copy, id);
            targetCaches.Add(copy);
        }
        if (addTargetCaches)
        {
            var anchor = target.Workbook.Root.Elements().FirstOrDefault(item =>
                item.Name.LocalName is "definedNames" or "calcPr" or "extLst");
            if (anchor == null)
                target.Workbook.Root.Add(targetCaches);
            else
                anchor.AddBeforeSelf(targetCaches);
        }
        output[target.WorkbookPart] = Save(target.Workbook);
        output[target.WorkbookRelationshipsPart] = Save(target.WorkbookRelationships);
    }

    /// <summary>
    /// 合并工作表级透视表关系和引用。
    /// </summary>
    private static void MergeWorksheets(ExcelPivotTemplateArchive output, PackageInfo source, PackageInfo target,
        ExcelPivotTemplateManifest manifest, CancellationToken token) =>
        MergeWorksheets(output, source, target, manifest.Entries, token);

    /// <summary>
    /// 合并指定透视表的工作表关系和引用。
    /// </summary>
    private static void MergeWorksheets(ExcelPivotTemplateArchive output, PackageInfo source, PackageInfo target,
        IEnumerable<ExcelPivotTemplateEntry> entries, CancellationToken token)
    {
        foreach (var group in entries.GroupBy(item => item.SheetName,
            StringComparer.OrdinalIgnoreCase))
        {
            token.ThrowIfCancellationRequested();
            if (!source.SheetParts.TryGetValue(group.Key, out var sourcePart)
                || !target.SheetParts.TryGetValue(group.Key, out var targetPart))
                throw new BingOfficesConfigurationException($"透视所属工作表未能映射: {group.Key}",
                    stage: BingOfficesStage.Write);
            if (!source.SheetRelationships.TryGetValue(sourcePart, out var sourceRelations))
                throw new InvalidDataException($"透视工作表关系缺失: {sourcePart}");
            if (!target.SheetDocuments.TryGetValue(targetPart, out var targetSheet))
                throw new InvalidDataException($"Provider 输出工作表部件缺失: {targetPart}");
            if (!target.SheetRelationships.TryGetValue(targetPart, out var targetRelations))
            {
                targetRelations = NewRelationships();
                target.SheetRelationships[targetPart] = targetRelations;
            }
            var pivotParts = FindElement(targetSheet.Root, "pivotTableParts")
                ?? new XElement(targetSheet.Root.Name.Namespace + "pivotTableParts");
            foreach (var entry in group)
            {
                token.ThrowIfCancellationRequested();
                var relation = FindRelationship(sourceRelations, entry.RelationshipId,
                    RelationshipPart(sourcePart));
                if (relation == null)
                    throw new InvalidDataException($"透视表关系缺失: {entry.RelationshipId}");
                var id = AddRelationship(EnsureRelationships(targetRelations), relation,
                    entry.PartName, targetPart);
                pivotParts.Add(new XElement(targetSheet.Root.Name.Namespace + "pivotTablePart",
                    new XAttribute(XName.Get("id", OfficeDocumentRelationshipNamespace), id)));
            }
            pivotParts.SetAttributeValue("count", pivotParts.Elements()
                .Count(item => item.Name.LocalName == "pivotTablePart"));
            if (pivotParts.Parent == null)
            {
                var anchor = targetSheet.Root.Elements().FirstOrDefault(item => item.Name.LocalName == "extLst");
                if (anchor == null)
                    targetSheet.Root.Add(pivotParts);
                else
                    anchor.AddBeforeSelf(pivotParts);
            }
            output[targetPart] = Save(targetSheet);
            output[RelationshipPart(targetPart)] = Save(targetRelations);
        }
    }

    /// <summary>
    /// 将 ZIP 部件索引复制到可按需读取的临时归档。
    /// </summary>
    private static ExcelPivotTemplateArchive ReadArchive(Stream source, CancellationToken token)
    {
        if (!source.CanSeek || !source.CanRead)
            throw new ArgumentException("透视模板合并要求可读取且可定位的流。", nameof(source));
        var position = source.Position;
        try
        {
            source.Position = 0;
            var staged = OpenPartsFile();
            try
            {
                Copy(source, staged, token);
                staged.Position = 0;
                var archive = new ZipArchive(staged, ZipArchiveMode.Read, leaveOpen: true);
                try
                {
                    var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
                    foreach (var entry in archive.Entries)
                    {
                        token.ThrowIfCancellationRequested();
                        if (entry.FullName.EndsWith("/", StringComparison.Ordinal))
                            continue;
                        var name = Normalize(entry.FullName);
                        if (!entries.TryAdd(name, entry))
                            throw new InvalidDataException($"ZIP 部件名称重复: {name}");
                    }
                    return new ExcelPivotTemplateArchive(staged, archive, entries, token);
                }
                catch
                {
                    archive.Dispose();
                    throw;
                }
            }
            catch
            {
                staged.Dispose();
                throw;
            }
        }
        finally
        {
            source.Position = position;
        }
    }

    /// <summary>
    /// 创建用于保存 ZIP 部件索引的临时文件。
    /// </summary>
    /// <returns>关闭时自动清理的异步文件流。</returns>
    private static FileStream OpenPartsFile() => new(Path.Combine(Path.GetTempPath(),
        "bing-offices-pivot-template-parts-" + Guid.NewGuid().ToString("N") + ".tmp"),
        FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous | FileOptions.SequentialScan);

    /// <summary>
    /// 创建供 Provider 加载的中间模板临时文件。
    /// </summary>
    /// <returns>关闭时自动清理的异步文件流。</returns>
    private static FileStream OpenStrippedTemplateFile() => new(Path.Combine(Path.GetTempPath(),
        "bing-offices-pivot-template-stripped-" + Guid.NewGuid().ToString("N") + ".tmp"),
        FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous | FileOptions.SequentialScan);

    /// <summary>
    /// 创建 Provider 生成工作簿使用的临时文件。
    /// </summary>
    /// <returns>关闭时自动清理的异步文件流。</returns>
    internal static FileStream OpenGeneratedFile() => new(Path.Combine(Path.GetTempPath(),
        "bing-offices-pivot-template-generated-" + Guid.NewGuid().ToString("N") + ".tmp"),
        FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous | FileOptions.SequentialScan);

    /// <summary>
    /// 创建 ClosedXML 异步桥接输出使用的外围临时文件。
    /// </summary>
    /// <returns>关闭时自动清理的异步文件流。</returns>
    internal static FileStream OpenAsyncOutputFile() => new(Path.Combine(Path.GetTempPath(),
        "bing-offices-pivot-template-async-" + Guid.NewGuid().ToString("N") + ".tmp"),
        FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920,
        FileOptions.DeleteOnClose | FileOptions.Asynchronous | FileOptions.SequentialScan);

    /// <summary>
    /// 预检 Provider 生成的 XLSX 包。
    /// </summary>
    /// <param name="source">Provider 生成的工作簿流。</param>
    /// <param name="token">取消令牌。</param>
    private void ValidateGeneratedPackage(Stream source, CancellationToken token)
    {
        try
        {
            ExcelXlsxZipPreflight.Validate(source, _snapshot.Limits, _provider, requireZip: true,
                cancellationToken: token, enforceSheetLimit: false, enforceCellLimits: false);
        }
        catch (BingOfficesResourceLimitException exception)
        {
            throw new BingOfficesResourceLimitException(exception.Message, exception, _provider,
                BingOfficesOperation.Export, BingOfficesStage.Write);
        }
        catch (BingOfficesImportException exception)
        {
            throw new BingOfficesExportException("Provider 生成的透视工作簿不是有效的 XLSX 包。",
                exception, _provider, BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 解析工作簿、工作表及其关系。
    /// </summary>
    private static PackageInfo ReadPackage(ExcelPivotTemplateArchive parts,
        CancellationToken token = default, long? maxXmlCharacters = null)
    {
        token.ThrowIfCancellationRequested();
        var rootRelations = LoadRequired(parts, "_rels/.rels", token, maxXmlCharacters);
        var root = rootRelations.Root.Elements().FirstOrDefault(item =>
            item.Name.LocalName == "Relationship" && ((string)item.Attribute("Type"))?.EndsWith(
                "/officeDocument", StringComparison.Ordinal) == true);
        if (root == null)
            throw new InvalidDataException("工作簿根关系缺失。");
        var workbookPart = ExcelPivotTemplateParts.ResolveTarget("_rels/.rels",
            (string)root.Attribute("Target"));
        var workbook = LoadRequired(parts, workbookPart, token, maxXmlCharacters);
        var workbookRelationshipsPart = RelationshipPart(workbookPart);
        var workbookRelationships = parts.TryGetValue(workbookRelationshipsPart, out var relBytes)
            ? Load(relBytes, token, maxXmlCharacters) : NewRelationships();
        var sheetParts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in workbook.Root.Descendants().Where(item => item.Name.LocalName == "sheet"))
        {
            var name = (string)sheet.Attribute("name");
            var id = GetRelationshipId(sheet);
            var relation = FindRelationship(workbookRelationships, id, workbookRelationshipsPart);
            if (string.IsNullOrWhiteSpace(name) || relation == null || relation.External)
                throw new InvalidDataException("工作簿 Sheet 关系无效。");
            sheetParts[name] = relation.PartName;
        }
        var sheetDocuments = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        var sheetRelationships = new Dictionary<string, XDocument>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in sheetParts.Values.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            token.ThrowIfCancellationRequested();
            sheetDocuments[part] = LoadRequired(parts, part, token, maxXmlCharacters);
            var relationshipPart = RelationshipPart(part);
            sheetRelationships[part] = parts.TryGetValue(relationshipPart, out var sheetRelBytes)
                ? Load(sheetRelBytes, token, maxXmlCharacters) : NewRelationships();
        }
        return new PackageInfo(workbookPart, workbookRelationshipsPart, workbook,
            workbookRelationships, sheetParts, sheetDocuments, sheetRelationships);
    }

    /// <summary>
    /// 从 XML 部件删除透视缓存声明。
    /// </summary>
    private static void RemoveWorkbookPivotCaches(XDocument document) =>
        FindElement(document.Root, "pivotCaches")?.Remove();

    /// <summary>
    /// 从 XML 部件删除透视表引用。
    /// </summary>
    private static void RemoveWorksheetPivotParts(XDocument document) =>
        FindElement(document.Root, "pivotTableParts")?.Remove();

    /// <summary>
    /// 从内容类型描述删除透视部件声明。
    /// </summary>
    private static void RemoveContentTypes(XDocument document, ISet<string> pivotParts)
    {
        foreach (var item in document.Root.Elements().Where(item => item.Name.LocalName == "Override").ToArray())
        {
            var part = Normalize(((string)item.Attribute("PartName"))?.TrimStart('/'));
            if (pivotParts.Contains(part))
                item.Remove();
        }
    }

    /// <summary>
    /// 从关系部件删除透视关系。
    /// </summary>
    private static void RemovePivotRelationships(XDocument document, ISet<string> pivotParts,
        string relationshipPart)
    {
        foreach (var item in document.Root.Elements().Where(item => item.Name.LocalName == "Relationship").ToArray())
        {
            var type = (string)item.Attribute("Type");
            var target = (string)item.Attribute("Target");
            var isPivot = type?.EndsWith("/pivotTable", StringComparison.Ordinal) == true
                || type?.EndsWith("/pivotCacheDefinition", StringComparison.Ordinal) == true
                || type?.EndsWith("/pivotCacheRecords", StringComparison.Ordinal) == true;
            var resolved = string.IsNullOrWhiteSpace(target) ? null :
                ExcelPivotTemplateParts.ResolveTarget(relationshipPart, target);
            if (isPivot || resolved != null && pivotParts.Contains(resolved))
                item.Remove();
        }
    }

    /// <summary>
    /// 复制普通部件到 ZIP。
    /// </summary>
    private static void AddEntry(ZipArchive archive, string name, byte[] bytes)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.Fastest);
        using var stream = entry.Open();
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// 合并全部已请求的透视更新。
    /// </summary>
    /// <param name="output">Provider 生成的部件集合。</param>
    /// <param name="generated">Provider 生成的工作簿流。</param>
    /// <param name="token">取消令牌。</param>
    private void MergeUpdates(ExcelPivotTemplateArchive output, Stream generated,
        CancellationToken token)
    {
        var generatedManifest = ExcelPivotTemplateManifest.Read(generated, _snapshot.Limits,
            _provider, token);
        ValidateGeneratedUpdateManifest(generatedManifest);
        var generatedEntries = ResolveGeneratedEntries(generatedManifest);
        MarkGeneratedCachesForRefreshOnLoad(output, UpdatedGeneratedCaches(generatedManifest), token,
            _snapshot.Limits.MaxXmlCharacters);

        var generatedPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
        PreservePivotLayout(output, generatedEntries, token);
        PrepareShrunkSources(output, generatedPackage, generatedEntries, token);
        foreach (var item in generatedEntries)
            RemoveSheet(output, generatedPackage, item.Entry.SheetName, token,
                _snapshot.Limits.MaxXmlCharacters);
        var targetPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
        PrepareUpdatedTargets(targetPackage, generatedEntries, token);
        foreach (var item in generatedEntries)
            MergeUpdatedWorksheet(output, generatedPackage, targetPackage, item.Entry,
                item.Update.Definition.TargetSheet, token);
    }

    /// <summary>
    /// 合并部分透视更新并保留未更新的模板部件。
    /// </summary>
    /// <param name="output">Provider 生成的部件集合。</param>
    /// <param name="generated">Provider 生成的工作簿流。</param>
    /// <param name="token">取消令牌。</param>
    /// <remarks>
    /// Provider 只处理请求中的透视，原模板中的其他透视、缓存记录和关系由桥接层恢复。
    /// 生成部件会先避开原模板路径和缓存标识，避免不同 Provider 的编号策略互相覆盖。
    /// </remarks>
    private void MergePartialUpdates(ExcelPivotTemplateArchive output, Stream generated,
        CancellationToken token)
    {
        var generatedManifest = ExcelPivotTemplateManifest.Read(generated, _snapshot.Limits,
            _provider, token);
        ValidateGeneratedUpdateManifest(generatedManifest);
        MarkGeneratedCachesForRefreshOnLoad(output, UpdatedGeneratedCaches(generatedManifest), token,
            _snapshot.Limits.MaxXmlCharacters);

        _snapshot.Source.Position = 0;
        using var original = ReadArchive(_snapshot.Source, token);
        var originalPackage = ReadPackage(original, token, _snapshot.Limits.MaxXmlCharacters);
        var preservedEntries = _snapshot.Plan.PreservedEntries;
        var preservedCaches = preservedEntries
            .Select(item => item.Cache)
            .GroupBy(item => item.Id)
            .Select(group => group.First())
            .ToArray();
        var preservedParts = SelectedPivotParts(preservedEntries, preservedCaches);
        foreach (var part in preservedParts.ToArray())
            preservedParts.Add(RelationshipPart(part));

        RenameGeneratedParts(output, generatedManifest, preservedParts, token,
            _snapshot.Limits.MaxXmlCharacters);
        generatedManifest = ReadGeneratedManifest(output, token);
        var cacheIds = CreateCacheIdMap(_snapshot.Plan.Manifest.Caches, generatedManifest.Caches);
        RemapGeneratedCacheIds(output, generatedManifest, cacheIds, token,
            _snapshot.Limits.MaxXmlCharacters);
        generatedManifest = ReadGeneratedManifest(output, token);

        var generatedPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
        var generatedEntries = ResolveGeneratedEntries(generatedManifest);
        PreservePivotLayout(output, generatedEntries, token);
        PrepareShrunkSources(output, generatedPackage, generatedEntries, token);
        foreach (var item in generatedEntries)
            RemoveSheet(output, generatedPackage, item.Entry.SheetName, token,
                _snapshot.Limits.MaxXmlCharacters);

        var targetPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
        CopyPivotPartsSubset(original, output, preservedEntries, preservedCaches, token);
        MergeContentTypesSubset(original, output, preservedEntries, preservedCaches, token,
            _snapshot.Limits.MaxXmlCharacters);
        MergeWorkbookSubset(output, originalPackage, targetPackage, preservedCaches, token);
        MergeWorksheets(output, originalPackage, targetPackage, preservedEntries, token);

        targetPackage = ReadPackage(output, token, _snapshot.Limits.MaxXmlCharacters);
        PrepareUpdatedTargets(targetPackage, generatedEntries, token);
        foreach (var item in generatedEntries)
            MergeUpdatedWorksheet(output, generatedPackage, targetPackage, item.Entry,
                item.Update.Definition.TargetSheet, token);
    }

    /// <summary>
    /// 校验 Provider 生成的更新部件数量和缓存关系。
    /// </summary>
    private void ValidateGeneratedUpdateManifest(ExcelPivotTemplateManifest manifest)
    {
        var expected = new HashSet<string>(_snapshot.Plan.NewPivots.Select(item => item.Name),
            StringComparer.OrdinalIgnoreCase);
        expected.UnionWith(_snapshot.Plan.Updates.Select(item => item.Entry.Name));
        if (manifest.Entries.Count != expected.Count
            || manifest.Entries.Any(item => !expected.Contains(item.Name)))
            throw UnsupportedUpdate("Provider 生成的透视与请求定义不一致。", BingOfficesStage.Write);
        var entries = manifest.Entries.Select(item => item.PartName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (manifest.Caches.Count == 0 || manifest.Caches.Any(cache => cache.Consumers.Count == 0
            || cache.Consumers.Any(consumer => !entries.Contains(consumer))))
            throw UnsupportedUpdate("Provider 生成了不属于本次请求的透视缓存消费者。", BingOfficesStage.Write);
    }

    /// <summary>
    /// 标记本次更新生成的缓存，以请求客户端在打开工作簿时刷新。
    /// </summary>
    /// <param name="output">Provider 生成的部件集合。</param>
    /// <param name="caches">本次输出的缓存描述。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="maxXmlCharacters">XML 部件允许的最大字符数。</param>
    /// <remarks>
    /// 该标记只写入新生成的缓存；模板中保留的未更新缓存保持原样。
    /// 标记表示客户端刷新请求，不代表 Provider 在服务端执行了缓存计算或保证客户端行为。
    /// </remarks>
    private static void MarkGeneratedCachesForRefreshOnLoad(
        ExcelPivotTemplateArchive output,
        IEnumerable<ExcelPivotTemplateCache> caches,
        CancellationToken token, long? maxXmlCharacters)
    {
        foreach (var cache in caches)
        {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(cache.DefinitionPart)
                || !output.TryGetValue(Normalize(cache.DefinitionPart), out var bytes))
                throw new InvalidDataException($"Provider 输出缺少生成的透视缓存定义: {cache.DefinitionPart}");
            var document = Load(bytes, token, maxXmlCharacters);
            if (document.Root == null || document.Root.Name.LocalName != "pivotCacheDefinition")
                throw new InvalidDataException($"生成的透视缓存定义根元素无效: {cache.DefinitionPart}");
            document.Root.SetAttributeValue("refreshOnLoad", "true");
            output[Normalize(cache.DefinitionPart)] = Save(document);
        }
    }

    /// <summary>
    /// 将 Provider 生成的透视名称映射到请求更新。
    /// </summary>
    private List<(ExcelPivotTemplateEntry Entry, ExcelPivotTemplateUpdate Update)> ResolveGeneratedEntries(
        ExcelPivotTemplateManifest manifest)
    {
        var updates = _snapshot.Plan.Updates.ToDictionary(item => item.Entry.Name,
            StringComparer.OrdinalIgnoreCase);
        var result = new List<(ExcelPivotTemplateEntry Entry, ExcelPivotTemplateUpdate Update)>(manifest.Entries.Count);
        foreach (var generatedEntry in manifest.Entries)
        {
            if (!updates.TryGetValue(generatedEntry.Name, out var update))
            {
                if (_snapshot.Plan.NewPivots.Any(pivot => string.Equals(pivot.Name,
                        generatedEntry.Name, StringComparison.OrdinalIgnoreCase)))
                    continue;
                throw UnsupportedUpdate("Provider 生成了请求之外的透视更新。", BingOfficesStage.Write);
            }
            result.Add((generatedEntry, update));
        }
        if (result.Count != updates.Count)
            throw UnsupportedUpdate("Provider 未生成全部请求的透视更新。", BingOfficesStage.Write);
        return result;
    }

    /// <summary>
    /// 选择至少被模板更新透视引用的本次生成缓存。
    /// </summary>
    private IEnumerable<ExcelPivotTemplateCache> UpdatedGeneratedCaches(
        ExcelPivotTemplateManifest manifest)
    {
        var updatedParts = manifest.Entries
            .Where(entry => _snapshot.Plan.Updates.Any(update => string.Equals(update.Entry.Name,
                entry.Name, StringComparison.OrdinalIgnoreCase)))
            .Select(entry => entry.PartName)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        return manifest.Caches.Where(cache => cache.Consumers.Any(updatedParts.Contains));
    }

    /// <summary>
    /// 读取当前部件集合的透视清单。
    /// </summary>
    private ExcelPivotTemplateManifest ReadGeneratedManifest(
        ExcelPivotTemplateArchive parts, CancellationToken token)
    {
        using var archive = CreateArchive(parts, token);
        return ExcelPivotTemplateManifest.Read(archive, _snapshot.Limits, _provider, token);
    }

    /// <summary>
    /// 获取指定透视和缓存所需的部件路径。
    /// </summary>
    private static HashSet<string> SelectedPivotParts(
        IEnumerable<ExcelPivotTemplateEntry> entries,
        IEnumerable<ExcelPivotTemplateCache> caches)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in entries)
            result.Add(Normalize(entry.PartName));
        foreach (var cache in caches)
        {
            if (!string.IsNullOrWhiteSpace(cache.DefinitionPart))
                result.Add(Normalize(cache.DefinitionPart));
            if (!string.IsNullOrWhiteSpace(cache.RecordsPart))
                result.Add(Normalize(cache.RecordsPart));
        }
        return result;
    }

    /// <summary>
    /// 生成不与原模板冲突的缓存标识映射。
    /// </summary>
    private static IReadOnlyDictionary<uint, uint> CreateCacheIdMap(
        IEnumerable<ExcelPivotTemplateCache> original,
        IEnumerable<ExcelPivotTemplateCache> generated)
    {
        var used = new HashSet<uint>(original.Select(item => item.Id));
        var result = new Dictionary<uint, uint>();
        var next = 1u;
        foreach (var cache in generated.OrderBy(item => item.Id))
        {
            while (used.Contains(next))
            {
                if (next == uint.MaxValue)
                    throw new InvalidDataException("透视缓存标识空间已耗尽。");
                next++;
            }
            result.Add(cache.Id, next);
            used.Add(next);
            if (next != uint.MaxValue)
                next++;
        }
        return result;
    }

    /// <summary>
    /// 将生成透视及工作簿声明中的缓存标识替换为新标识。
    /// </summary>
    private static void RemapGeneratedCacheIds(ExcelPivotTemplateArchive output,
        ExcelPivotTemplateManifest manifest, IReadOnlyDictionary<uint, uint> cacheIds,
        CancellationToken token, long? maxXmlCharacters)
    {
        if (cacheIds.Count == 0)
            return;
        token.ThrowIfCancellationRequested();
        var workbook = LoadRequired(output, manifest.WorkbookPart, token, maxXmlCharacters);
        var caches = FindElement(workbook.Root, "pivotCaches")
            ?? throw new InvalidDataException("Provider 输出缺少透视缓存声明。");
        foreach (var cache in caches.Elements().Where(item => item.Name.LocalName == "pivotCache"))
        {
            token.ThrowIfCancellationRequested();
            var attribute = cache.Attribute("cacheId");
            if (attribute == null || !uint.TryParse(attribute.Value, out var oldId)
                || !cacheIds.TryGetValue(oldId, out var newId))
                throw new InvalidDataException("Provider 输出的透视缓存标识无法映射。");
            attribute.Value = newId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        output[manifest.WorkbookPart] = Save(workbook);

        foreach (var entry in manifest.Entries)
        {
            token.ThrowIfCancellationRequested();
            var document = LoadRequired(output, entry.PartName, token, maxXmlCharacters);
            var attribute = document.Root?.Attributes()
                .FirstOrDefault(item => item.Name.LocalName == "cacheId");
            if (attribute == null || !uint.TryParse(attribute.Value, out var oldId)
                || !cacheIds.TryGetValue(oldId, out var newId))
                throw new InvalidDataException("Provider 输出的透视定义缓存标识无法映射。");
            attribute.Value = newId.ToString(System.Globalization.CultureInfo.InvariantCulture);
            output[entry.PartName] = Save(document);
        }
    }

    /// <summary>
    /// 避开原模板部件路径，防止 Provider 生成部件覆盖未更新透视。
    /// </summary>
    private static void RenameGeneratedParts(ExcelPivotTemplateArchive output,
        ExcelPivotTemplateManifest generatedManifest, ISet<string> reservedParts,
        CancellationToken token, long? maxXmlCharacters)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var sequence = 1;
        foreach (var part in generatedManifest.Parts.Select(Normalize))
        {
            token.ThrowIfCancellationRequested();
            if (!output.ContainsKey(part))
                throw new InvalidDataException($"Provider 输出缺少透视部件: {part}");
            var relationPart = RelationshipPart(part);
            var candidate = part;
            while (reservedParts.Contains(candidate)
                || output.ContainsKey(candidate) && !string.Equals(candidate, part,
                    StringComparison.OrdinalIgnoreCase)
                || map.Values.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                candidate = GeneratedPartName(part, sequence++);
            var candidateRelations = RelationshipPart(candidate);
            if (output.ContainsKey(relationPart)
                && (reservedParts.Contains(candidateRelations)
                    || output.ContainsKey(candidateRelations) && !string.Equals(candidateRelations,
                        relationPart, StringComparison.OrdinalIgnoreCase)
                    || map.Values.Contains(candidateRelations, StringComparer.OrdinalIgnoreCase)))
            {
                do
                {
                    candidate = GeneratedPartName(part, sequence++);
                    candidateRelations = RelationshipPart(candidate);
                }
                while (reservedParts.Contains(candidate) || output.ContainsKey(candidate)
                    || map.Values.Contains(candidate, StringComparer.OrdinalIgnoreCase)
                    || reservedParts.Contains(candidateRelations) || output.ContainsKey(candidateRelations)
                    || map.Values.Contains(candidateRelations, StringComparer.OrdinalIgnoreCase));
            }
            if (!string.Equals(candidate, part, StringComparison.OrdinalIgnoreCase))
            {
                map[part] = candidate;
                if (output.ContainsKey(relationPart))
                    map[relationPart] = candidateRelations;
            }
        }
        if (map.Count == 0)
            return;
        foreach (var item in map.ToArray())
        {
            token.ThrowIfCancellationRequested();
            if (!output.TryGetValue(item.Key, out var bytes))
                throw new InvalidDataException($"Provider 输出缺少待重命名部件: {item.Key}");
            output.Remove(item.Key);
            output[item.Value] = bytes;
        }

        if (output.TryGetValue("[Content_Types].xml", out var contentTypes))
        {
            var document = Load(contentTypes, token, maxXmlCharacters);
            foreach (var item in document.Root.Elements().Where(item => item.Name.LocalName == "Override"))
            {
                token.ThrowIfCancellationRequested();
                var part = Normalize(((string)item.Attribute("PartName"))?.TrimStart('/'));
                if (map.TryGetValue(part, out var mapped))
                    item.SetAttributeValue("PartName", "/" + mapped);
            }
            output["[Content_Types].xml"] = Save(document);
        }

        foreach (var relationName in output.Keys
                     .Where(item => item.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                     .ToArray())
        {
            token.ThrowIfCancellationRequested();
            var document = Load(output[relationName], token, maxXmlCharacters);
            var changed = false;
            foreach (var relation in document.Root.Elements()
                         .Where(item => item.Name.LocalName == "Relationship"))
            {
                token.ThrowIfCancellationRequested();
                if (string.Equals((string)relation.Attribute("TargetMode"), "External",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                var target = (string)relation.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target))
                    continue;
                var resolved = ExcelPivotTemplateParts.ResolveTarget(relationName, target);
                if (map.TryGetValue(resolved, out var mapped))
                {
                    relation.SetAttributeValue("Target", RelativeTarget(RelationshipOwner(relationName), mapped));
                    changed = true;
                }
            }
            if (changed)
                output[relationName] = Save(document);
        }
    }

    /// <summary>
    /// 生成唯一的 Provider 透视部件名称。
    /// </summary>
    private static string GeneratedPartName(string part, int sequence)
    {
        var slash = part.LastIndexOf('/');
        var extension = Path.GetExtension(part);
        var folder = slash < 0 ? string.Empty : part[..(slash + 1)];
        return folder + "__BingPivotUpdate_"
            + sequence.ToString("000", System.Globalization.CultureInfo.InvariantCulture) + extension;
    }

    /// <summary>
    /// 复制指定的原始透视部件及其关系文件。
    /// </summary>
    private static void CopyPivotPartsSubset(ExcelPivotTemplateArchive original,
        ExcelPivotTemplateArchive output, IEnumerable<ExcelPivotTemplateEntry> entries,
        IEnumerable<ExcelPivotTemplateCache> caches, CancellationToken token)
    {
        var names = SelectedPivotParts(entries, caches);
        foreach (var name in names.ToArray())
            names.Add(RelationshipPart(name));
        foreach (var name in names)
        {
            token.ThrowIfCancellationRequested();
            if (!original.TryGetValue(name, out var bytes))
                continue;
            if (output.ContainsKey(name))
                throw new InvalidDataException($"Provider 输出已占用透视部件路径: {name}");
            output[name] = bytes;
        }
    }

    /// <summary>
    /// 合并指定原始透视部件的内容类型声明。
    /// </summary>
    private static void MergeContentTypesSubset(ExcelPivotTemplateArchive original,
        ExcelPivotTemplateArchive output, IEnumerable<ExcelPivotTemplateEntry> entries,
        IEnumerable<ExcelPivotTemplateCache> caches, CancellationToken token,
        long? maxXmlCharacters)
    {
        if (!original.TryGetValue("[Content_Types].xml", out var originalBytes)
            || !output.TryGetValue("[Content_Types].xml", out var outputBytes))
            throw new InvalidDataException("透视模板缺少内容类型部件。");
        var source = Load(originalBytes, token, maxXmlCharacters);
        var target = Load(outputBytes, token, maxXmlCharacters);
        var pivotParts = SelectedPivotParts(entries, caches);
        var existing = target.Root.Elements().Where(item => item.Name.LocalName == "Override")
            .Select(item => Normalize(((string)item.Attribute("PartName"))?.TrimStart('/')))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var item in source.Root.Elements().Where(item => item.Name.LocalName == "Override"))
        {
            token.ThrowIfCancellationRequested();
            var part = Normalize(((string)item.Attribute("PartName"))?.TrimStart('/'));
            if (pivotParts.Contains(part) && existing.Add(part))
                target.Root.Add(new XElement(item));
        }
        output["[Content_Types].xml"] = Save(target);
    }

    /// <summary>
    /// 将指定原始缓存追加到 Provider 输出的工作簿缓存集合。
    /// </summary>
    private static void MergeWorkbookSubset(ExcelPivotTemplateArchive output, PackageInfo source,
        PackageInfo target, IEnumerable<ExcelPivotTemplateCache> caches, CancellationToken token)
    {
        var selected = caches.ToArray();
        if (selected.Length == 0)
            return;
        var originalCaches = FindElement(source.Workbook.Root, "pivotCaches")
            ?? throw new InvalidDataException("透视清单缺少工作簿缓存声明。");
        var targetCaches = FindElement(target.Workbook.Root, "pivotCaches");
        if (targetCaches == null)
        {
            targetCaches = new XElement(target.Workbook.Root.Name.Namespace + "pivotCaches");
            var anchor = target.Workbook.Root.Elements().FirstOrDefault(item =>
                item.Name.LocalName is "definedNames" or "calcPr" or "extLst");
            if (anchor == null)
                target.Workbook.Root.Add(targetCaches);
            else
                anchor.AddBeforeSelf(targetCaches);
        }
        var existingIds = targetCaches.Elements().Where(item => item.Name.LocalName == "pivotCache")
            .Select(item => (string)item.Attribute("cacheId"))
            .Where(item => uint.TryParse(item, out _))
            .Select(item => uint.Parse(item, System.Globalization.CultureInfo.InvariantCulture))
            .ToHashSet();
        var targetRelations = EnsureRelationships(target.WorkbookRelationships);
        foreach (var cache in selected)
        {
            token.ThrowIfCancellationRequested();
            var originalCache = originalCaches.Elements()
                .FirstOrDefault(item => item.Name.LocalName == "pivotCache"
                    && uint.TryParse((string)item.Attribute("cacheId"), out var id) && id == cache.Id);
            if (originalCache == null)
                throw new InvalidDataException($"透视缓存声明缺少 cacheId: {cache.Id}");
            if (!existingIds.Add(cache.Id))
                throw new InvalidDataException($"透视缓存标识冲突: {cache.Id}");
            var relation = FindRelationship(source.WorkbookRelationships, cache.RelationshipId,
                source.WorkbookRelationshipsPart);
            if (relation == null)
                throw new InvalidDataException($"透视缓存关系缺失: {cache.RelationshipId}");
            var id = AddRelationship(targetRelations, relation, cache.DefinitionPart, target.WorkbookPart);
            var copy = new XElement(originalCache);
            SetRelationshipId(copy, id);
            targetCaches.Add(copy);
        }
        output[target.WorkbookPart] = Save(target.Workbook);
        output[target.WorkbookRelationshipsPart] = Save(target.WorkbookRelationships);
    }

    /// <summary>
    /// 从 Provider 输出中移除更新用临时工作表。
    /// </summary>
    private static void RemoveSheet(ExcelPivotTemplateArchive output, PackageInfo package,
        string sheetName, CancellationToken token, long? maxXmlCharacters)
    {
        if (!package.SheetParts.TryGetValue(sheetName, out var sheetPart))
            throw new InvalidDataException($"透视更新临时工作表缺失: {sheetName}");
        var sheet = package.Workbook.Root.Descendants().FirstOrDefault(item =>
            item.Name.LocalName == "sheet" && string.Equals((string)item.Attribute("name"),
                sheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet == null)
            throw new InvalidDataException($"透视更新工作簿声明缺失: {sheetName}");
        var relationshipId = GetRelationshipId(sheet);
        sheet.Remove();
        var relation = package.WorkbookRelationships.Root.Elements().FirstOrDefault(item =>
            item.Name.LocalName == "Relationship"
            && string.Equals((string)item.Attribute("Id"), relationshipId, StringComparison.Ordinal));
        relation?.Remove();
        output[package.WorkbookPart] = Save(package.Workbook);
        output[package.WorkbookRelationshipsPart] = Save(package.WorkbookRelationships);
        output.Remove(sheetPart);
        output.Remove(RelationshipPart(sheetPart));
        if (output.TryGetValue("[Content_Types].xml", out var contentTypes))
        {
            var document = Load(contentTypes, token, maxXmlCharacters);
            foreach (var item in document.Root.Elements().Where(item => item.Name.LocalName == "Override")
                         .ToArray())
            {
                token.ThrowIfCancellationRequested();
                var part = Normalize(((string)item.Attribute("PartName"))?.TrimStart('/'));
                if (string.Equals(part, sheetPart, StringComparison.OrdinalIgnoreCase))
                    item.Remove();
            }
            output["[Content_Types].xml"] = Save(document);
        }
    }

    /// <summary>
    /// 将临时工作表上的透视关系挂到原目标工作表。
    /// </summary>
    private static void MergeUpdatedWorksheet(ExcelPivotTemplateArchive output, PackageInfo source,
        PackageInfo target, ExcelPivotTemplateEntry entry, string targetSheetName,
        CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!target.SheetParts.TryGetValue(targetSheetName, out var targetPart)
            || !target.SheetDocuments.TryGetValue(targetPart, out var targetSheet))
            throw new BingOfficesConfigurationException($"透视更新目标工作表不存在: {targetSheetName}",
                stage: BingOfficesStage.Write);
        if (!source.SheetParts.TryGetValue(entry.SheetName, out var sourcePart)
            || !source.SheetRelationships.TryGetValue(sourcePart, out var sourceRelations))
            throw new InvalidDataException("透视更新临时工作表关系缺失。");
        var relation = FindRelationship(sourceRelations, entry.RelationshipId,
            RelationshipPart(sourcePart));
        if (relation == null)
            throw new InvalidDataException("透视更新关系缺失。");
        if (!target.SheetRelationships.TryGetValue(targetPart, out var targetRelations))
        {
            targetRelations = NewRelationships();
            target.SheetRelationships[targetPart] = targetRelations;
        }
        var pivotParts = FindElement(targetSheet.Root, "pivotTableParts")
            ?? new XElement(targetSheet.Root.Name.Namespace + "pivotTableParts");
        var id = AddRelationship(EnsureRelationships(targetRelations), relation, entry.PartName,
            targetPart);
        pivotParts.Add(new XElement(targetSheet.Root.Name.Namespace + "pivotTablePart",
            new XAttribute(XName.Get("id", OfficeDocumentRelationshipNamespace), id)));
        pivotParts.SetAttributeValue("count", pivotParts.Elements()
            .Count(item => item.Name.LocalName == "pivotTablePart"));
        if (pivotParts.Parent == null)
            targetSheet.Root.Add(pivotParts);
        output[targetPart] = Save(targetSheet);
        output[RelationshipPart(targetPart)] = Save(targetRelations);
    }

    /// <summary>
    /// 校验更新区域并清除旧透视正文。
    /// </summary>
    private void PrepareUpdatedTargets(PackageInfo target,
        IReadOnlyList<(ExcelPivotTemplateEntry Entry, ExcelPivotTemplateUpdate Update)> updates,
        CancellationToken token)
    {
        var updatedNames = new HashSet<string>(updates.Select(item => item.Entry.Name),
            StringComparer.OrdinalIgnoreCase);
        var targets = new List<(string SheetName, XDocument Sheet, ExcelRangeDefinition Previous,
            ExcelRangeDefinition Current)>();
        foreach (var item in updates)
        {
            token.ThrowIfCancellationRequested();
            var previous = ParsePivotRange(item.Update.Entry.TargetReference);
            var current = ParsePivotRange(item.Entry.TargetReference);
            foreach (var other in _snapshot.Plan.Manifest.Entries)
            {
                token.ThrowIfCancellationRequested();
                if (updatedNames.Contains(other.Name))
                    continue;
                if (string.Equals(other.SheetName, item.Update.Definition.TargetSheet,
                        StringComparison.OrdinalIgnoreCase)
                    && Overlaps(current, ParsePivotRange(other.TargetReference)))
                    throw TargetConflict(item.Entry.Name, "更新区域与其他模板透视区域重叠。");
            }
            foreach (var other in updates)
            {
                token.ThrowIfCancellationRequested();
                if (ReferenceEquals(item.Entry, other.Entry)
                    || !string.Equals(item.Update.Definition.TargetSheet,
                        other.Update.Definition.TargetSheet, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (Overlaps(current, ParsePivotRange(other.Entry.TargetReference)))
                    throw TargetConflict(item.Entry.Name, "更新区域与另一项透视更新重叠。");
            }
            if (!target.SheetParts.TryGetValue(item.Update.Definition.TargetSheet, out var part)
                || !target.SheetDocuments.TryGetValue(part, out var sheet))
                throw new BingOfficesConfigurationException(
                    $"透视更新目标工作表不存在: {item.Update.Definition.TargetSheet}",
                    stage: BingOfficesStage.Preflight);
            var occupied = FindOccupiedCellOutside(sheet, current, previous, token);
            if (occupied != null)
                throw TargetConflict(item.Entry.Name,
                    $"更新区域覆盖了原透视区域外的非空单元格: {occupied}");
            targets.Add((item.Update.Definition.TargetSheet, sheet, previous, current));
        }

        foreach (var targetInfo in targets)
            ClearPivotValues(targetInfo.Sheet, targetInfo.Previous, token);
    }

    /// <summary>
    /// 清除收缩后不再属于透视来源的旧单元格值。
    /// </summary>
    private void PrepareShrunkSources(ExcelPivotTemplateArchive output, PackageInfo generated,
        IReadOnlyList<(ExcelPivotTemplateEntry Entry, ExcelPivotTemplateUpdate Update)> updates,
        CancellationToken token)
    {
        foreach (var item in updates)
        {
            token.ThrowIfCancellationRequested();
            var cache = item.Update.Entry.Cache;
            var definition = item.Update.Definition;
            if (!string.Equals(cache.SourceType, "worksheet", StringComparison.OrdinalIgnoreCase)
                || !string.IsNullOrWhiteSpace(cache.SourceName)
                || !string.Equals(cache.SourceSheet, definition.SourceSheet,
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var previous = ParsePivotRange(cache.SourceReference);
            var current = new ExcelRangeDefinition
            {
                StartRow = definition.SourceRange.StartRow,
                StartColumn = definition.SourceRange.StartColumn,
                EndRow = definition.SourceRange.EndRow,
                EndColumn = definition.SourceRange.EndColumn
            };
            if (previous.StartRow != current.StartRow || previous.StartColumn != current.StartColumn
                || Contains(current, previous) || !Contains(previous, current))
                continue;

            var retainedSources = _snapshot.Plan.PreservedEntries
                .Select(entry => entry.Cache)
                .Where(other => string.Equals(other.SourceType, "worksheet", StringComparison.OrdinalIgnoreCase)
                    && string.IsNullOrWhiteSpace(other.SourceName)
                    && string.Equals(other.SourceSheet, cache.SourceSheet,
                        StringComparison.OrdinalIgnoreCase))
                .Select(other => ParsePivotRange(other.SourceReference));
            var requestedSources = updates.Where(other => !ReferenceEquals(other.Entry, item.Entry)
                    && string.Equals(other.Update.Definition.SourceSheet, definition.SourceSheet,
                        StringComparison.OrdinalIgnoreCase))
                .Select(other => new ExcelRangeDefinition
                {
                    StartRow = other.Update.Definition.SourceRange.StartRow,
                    StartColumn = other.Update.Definition.SourceRange.StartColumn,
                    EndRow = other.Update.Definition.SourceRange.EndRow,
                    EndColumn = other.Update.Definition.SourceRange.EndColumn
                })
                .Concat(_snapshot.Plan.NewPivots.Where(pivot => string.Equals(pivot.SourceSheet,
                        definition.SourceSheet, StringComparison.OrdinalIgnoreCase))
                    .Select(pivot => new ExcelRangeDefinition
                    {
                        StartRow = pivot.SourceRange.StartRow,
                        StartColumn = pivot.SourceRange.StartColumn,
                        EndRow = pivot.SourceRange.EndRow,
                        EndColumn = pivot.SourceRange.EndColumn
                    }));
            if (retainedSources.Concat(requestedSources)
                .Any(range => Overlaps(previous, range) && !Contains(current, range)))
                continue;

            if (!generated.SheetParts.TryGetValue(definition.SourceSheet, out var part)
                || !generated.SheetDocuments.TryGetValue(part, out var sheet))
                continue;
            ClearRangeOutside(sheet, previous, current, token);
            output[part] = Save(sheet);
        }
    }

    /// <summary>
    /// 保留更新定义未覆盖的透视布局和样式。
    /// </summary>
    private void PreservePivotLayout(ExcelPivotTemplateArchive output,
        IReadOnlyList<(ExcelPivotTemplateEntry Entry, ExcelPivotTemplateUpdate Update)> updates,
        CancellationToken token)
    {
        if (updates.Count == 0)
            return;
        _snapshot.Source.Position = 0;
        using var original = ReadArchive(_snapshot.Source, token);
        foreach (var item in updates)
        {
            token.ThrowIfCancellationRequested();
            if (!original.TryGetValue(item.Update.Entry.PartName, out var originalBytes)
                || !output.TryGetValue(item.Entry.PartName, out var generatedBytes))
                throw new InvalidDataException($"透视布局部件缺失: {item.Entry.Name}");

            var originalDocument = Load(originalBytes, token, _snapshot.Limits.MaxXmlCharacters);
            var generatedDocument = Load(generatedBytes, token, _snapshot.Limits.MaxXmlCharacters);
            var originalRoot = originalDocument.Root;
            var generatedRoot = generatedDocument.Root;
            if (originalRoot == null || generatedRoot == null
                || originalRoot.Name.LocalName != "pivotTableDefinition"
                || generatedRoot.Name.LocalName != "pivotTableDefinition")
                throw new InvalidDataException($"透视布局定义无效: {item.Entry.Name}");

            foreach (var name in PivotLayoutAttributes)
            {
                token.ThrowIfCancellationRequested();
                var value = originalRoot.Attribute(name);
                if (value != null)
                    generatedRoot.SetAttributeValue(name, value.Value);
            }

            var hasFieldIndexedFormatting = PivotLayoutElements.Any(name =>
                FindElement(originalRoot, name) != null);
            if (hasFieldIndexedFormatting && !item.Update.Entry.Cache.Fields.SequenceEqual(
                    item.Entry.Cache.Fields, StringComparer.OrdinalIgnoreCase))
                throw UnsupportedUpdate("模板透视包含字段索引格式且更新改变了来源字段顺序。",
                    BingOfficesStage.Write);
            foreach (var name in PivotLayoutElements)
            {
                token.ThrowIfCancellationRequested();
                CopyPivotLayoutElement(originalRoot, generatedRoot, name);
            }

            var originalStyle = FindElement(originalRoot, "pivotTableStyleInfo");
            if (originalStyle != null)
            {
                CopyPivotLayoutElement(originalRoot, generatedRoot, "pivotTableStyleInfo");
            }

            output[item.Entry.PartName] = Save(generatedDocument);
        }
    }

    /// <summary>
    /// 将模板中的布局元素按工作表格式顺序复制到更新定义。
    /// </summary>
    private static void CopyPivotLayoutElement(XElement source, XElement target, string name)
    {
        var sourceElement = FindElement(source, name);
        if (sourceElement == null)
            return;
        var targetElement = FindElement(target, name);
        var copy = new XElement(sourceElement);
        if (targetElement != null)
        {
            targetElement.ReplaceWith(copy);
            return;
        }

        var order = PivotDefinitionElementOrder(name);
        var anchor = target.Elements().FirstOrDefault(element =>
            PivotDefinitionElementOrder(element.Name.LocalName) > order);
        if (anchor == null)
            target.Add(copy);
        else
            anchor.AddBeforeSelf(copy);
    }

    /// <summary>
    /// 获取透视定义子元素的工作表格式顺序。
    /// </summary>
    private static int PivotDefinitionElementOrder(string name) => name switch
    {
        "location" => 0,
        "pivotFields" => 1,
        "rowFields" => 2,
        "rowItems" => 3,
        "colFields" => 4,
        "colItems" => 5,
        "pageFields" => 6,
        "dataFields" => 7,
        "formats" => 8,
        "conditionalFormats" => 9,
        "chartFormats" => 10,
        "pivotHierarchies" => 11,
        "pivotTableStyleInfo" => 12,
        "filters" => 13,
        "rowHierarchiesUsage" => 14,
        "colHierarchiesUsage" => 15,
        "extLst" => 16,
        _ => int.MaxValue
    };

    /// <summary>
    /// 清除旧范围中未被新范围覆盖的值和公式。
    /// </summary>
    private static void ClearRangeOutside(XDocument sheet, ExcelRangeDefinition previous,
        ExcelRangeDefinition current, CancellationToken token)
    {
        foreach (var cell in sheet.Descendants().Where(item => item.Name.LocalName == "c").ToArray())
        {
            token.ThrowIfCancellationRequested();
            var address = ParsePivotRange((string)cell.Attribute("r"));
            if (!Contains(previous, address) || Contains(current, address))
                continue;
            foreach (var value in cell.Elements().Where(item => item.Name.LocalName is "v" or "f" or "is").ToArray())
                value.Remove();
            cell.Attribute("t")?.Remove();
        }
    }

    /// <summary>
    /// 判断新区域是否覆盖旧透视区域以外的非空单元格。
    /// </summary>
    private static string FindOccupiedCellOutside(XDocument sheet, ExcelRangeDefinition current,
        ExcelRangeDefinition previous, CancellationToken token)
    {
        foreach (var cell in sheet.Descendants().Where(item => item.Name.LocalName == "c"))
        {
            token.ThrowIfCancellationRequested();
            if (!HasCellValue(cell))
                continue;
            var address = ParsePivotRange((string)cell.Attribute("r"));
            if (Contains(current, address) && !Contains(previous, address))
                return (string)cell.Attribute("r");
        }
        return null;
    }

    /// <summary>
    /// 移除旧透视区域内的值和公式，同时保留单元格样式。
    /// </summary>
    private static void ClearPivotValues(XDocument sheet, ExcelRangeDefinition range,
        CancellationToken token)
    {
        foreach (var cell in sheet.Descendants().Where(item => item.Name.LocalName == "c").ToArray())
        {
            token.ThrowIfCancellationRequested();
            if (!Contains(range, ParsePivotRange((string)cell.Attribute("r"))))
                continue;
            foreach (var value in cell.Elements().Where(item => item.Name.LocalName is "v" or "f" or "is").ToArray())
                value.Remove();
            cell.Attribute("t")?.Remove();
        }
    }

    /// <summary>
    /// 判断单元格是否包含值、公式或内联文本。
    /// </summary>
    private static bool HasCellValue(XElement cell) =>
        cell.Elements().Any(item => item.Name.LocalName is "v" or "f" or "is");

    /// <summary>
    /// 解析 XLSX 单元格或区域引用。
    /// </summary>
    private static ExcelRangeDefinition ParsePivotRange(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            throw new InvalidDataException("透视目标区域为空。");
        var parts = reference.Split(':');
        if (parts.Length is < 1 or > 2 || !TryParseCell(parts[0], out var startRow, out var startColumn))
            throw new InvalidDataException($"透视目标区域无效: {reference}");
        var endRow = startRow;
        var endColumn = startColumn;
        if (parts.Length == 2 && !TryParseCell(parts[1], out endRow, out endColumn))
            throw new InvalidDataException($"透视目标区域无效: {reference}");
        if (endRow < startRow || endColumn < startColumn)
            throw new InvalidDataException($"透视目标区域无效: {reference}");
        return new ExcelRangeDefinition
        {
            StartRow = startRow,
            StartColumn = startColumn,
            EndRow = endRow,
            EndColumn = endColumn
        };
    }

    /// <summary>
    /// 解析绝对或相对 XLSX 单元格引用。
    /// </summary>
    private static bool TryParseCell(string reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
            return false;
        var index = 0;
        if (reference[index] == '$')
            index++;
        var hasColumn = false;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            var character = char.ToUpperInvariant(reference[index++]);
            if (character is < 'A' or > 'Z')
                return false;
            column = checked(column * 26 + character - 'A' + 1);
            if (column > 16384)
                return false;
            hasColumn = true;
        }
        if (!hasColumn || index < reference.Length && reference[index] == '$')
        {
            if (index < reference.Length && reference[index] == '$')
                index++;
        }
        var digitStart = index;
        while (index < reference.Length && reference[index] is >= '0' and <= '9')
        {
            row = checked(row * 10 + reference[index++] - '0');
            if (row > 1048576)
                return false;
        }
        if (!hasColumn || digitStart == index || index != reference.Length
            || row is <= 0 or > 1048576 || column is <= 0 or > 16384)
            return false;
        row--;
        column--;
        return true;
    }

    /// <summary>
    /// 判断两个矩形区域是否相交。
    /// </summary>
    private static bool Overlaps(ExcelRangeDefinition left, ExcelRangeDefinition right) =>
        left.StartRow <= right.EndRow && right.StartRow <= left.EndRow
        && left.StartColumn <= right.EndColumn && right.StartColumn <= left.EndColumn;

    /// <summary>
    /// 判断外层区域是否完整包含内层区域。
    /// </summary>
    private static bool Contains(ExcelRangeDefinition outer, ExcelRangeDefinition inner) =>
        outer.StartRow <= inner.StartRow && outer.StartColumn <= inner.StartColumn
        && outer.EndRow >= inner.EndRow && outer.EndColumn >= inner.EndColumn;

    /// <summary>
    /// 创建透视目标区域冲突异常。
    /// </summary>
    private BingOfficesConfigurationException TargetConflict(string pivotName, string reason) =>
        new($"{_provider}: 透视更新目标冲突 ({pivotName})：{reason}",
            stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 创建当前部件集合对应的临时 ZIP 文件流。
    /// </summary>
    /// <param name="parts">待写入的部件集合。</param>
    /// <param name="token">取消令牌。</param>
    /// <param name="compressionLevel">目标压缩级别。</param>
    /// <returns>已定位到文件开头且关闭时自动清理的 ZIP 文件流。</returns>
    private static FileStream CreateArchive(ExcelPivotTemplateArchive parts,
        CancellationToken token, CompressionLevel compressionLevel = CompressionLevel.Fastest)
    {
        var path = Path.Combine(Path.GetTempPath(),
            "bing-offices-pivot-template-archive-" + Guid.NewGuid().ToString("N") + ".tmp");
        var result = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
            81920, FileOptions.DeleteOnClose | FileOptions.Asynchronous | FileOptions.SequentialScan);
        try
        {
            using (var archive = new ZipArchive(result, ZipArchiveMode.Create, leaveOpen: true))
                parts.WriteTo(archive, token, compressionLevel);
            result.Position = 0;
            return result;
        }
        catch
        {
            result.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 创建模板更新不支持异常。
    /// </summary>
    private BingOfficesUnsupportedFeatureException UnsupportedUpdate(string message,
        BingOfficesStage stage) => new($"{_provider}: {message}", provider: _provider,
        operation: BingOfficesOperation.Export, stage: stage);

    /// <summary>
    /// 复制流并检查取消。
    /// </summary>
    private static void Copy(Stream source, Stream destination, CancellationToken token)
    {
        var buffer = new byte[81920];
        int count;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            count = source.Read(buffer, 0, buffer.Length);
            if (count <= 0)
                break;
            destination.Write(buffer, 0, count);
        }
        token.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 读取 ZIP 条目。
    /// </summary>
    private static byte[] ReadEntry(ZipArchiveEntry entry, CancellationToken token = default)
    {
        using var input = entry.Open();
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        int count;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            count = input.Read(buffer, 0, buffer.Length);
            if (count <= 0)
                break;
            output.Write(buffer, 0, count);
        }
        token.ThrowIfCancellationRequested();
        return output.ToArray();
    }

    /// <summary>
    /// 编辑 XML ZIP 条目。
    /// </summary>
    private static byte[] EditXml(ZipArchiveEntry entry, Action<XDocument> edit,
        CancellationToken token = default, long? maxXmlCharacters = null)
    {
        var document = Load(ReadEntry(entry, token), token, maxXmlCharacters);
        edit(document);
        return Save(document);
    }

    /// <summary>
    /// 加载 XML 部件。
    /// </summary>
    private static XDocument Load(byte[] bytes, CancellationToken token = default,
        long? maxXmlCharacters = null)
    {
        using var stream = new CancellationReadStream(new MemoryStream(bytes, writable: false), token);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersFromEntities = 0,
            MaxCharactersInDocument = maxXmlCharacters ?? MaximumXmlCharacters
        });
        token.ThrowIfCancellationRequested();
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    /// <summary>
    /// 加载必需 XML 部件。
    /// </summary>
    private static XDocument LoadRequired(ExcelPivotTemplateArchive parts, string name,
        CancellationToken token = default, long? maxXmlCharacters = null) =>
        parts.TryGetValue(name, out var bytes)
            ? Load(bytes, token, maxXmlCharacters)
            : throw new InvalidDataException($"输出缺少 XML 部件: {name}");

    /// <summary>
    /// 保存 XML 部件。
    /// </summary>
    private static byte[] Save(XDocument document)
    {
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            Indent = false
        }))
            document.Save(writer);
        return stream.ToArray();
    }

    /// <summary>
    /// 获取 XML 中指定本地名称的第一个元素。
    /// </summary>
    private static XElement FindElement(XElement root, string localName) =>
        root?.Elements().FirstOrDefault(item => item.Name.LocalName == localName);

    /// <summary>
    /// 获取关系标识属性。
    /// </summary>
    private static string GetRelationshipId(XElement element) =>
        element.Attributes().FirstOrDefault(item => item.Name.LocalName == "id")?.Value;

    /// <summary>
    /// 设置关系标识属性。
    /// </summary>
    private static void SetRelationshipId(XElement element, string id)
    {
        var existing = element.Attributes().FirstOrDefault(item => item.Name.LocalName == "id");
        element.SetAttributeValue(existing?.Name ?? XName.Get("id", OfficeDocumentRelationshipNamespace), id);
    }

    /// <summary>
    /// 查找关系。
    /// </summary>
    private static RelationshipInfo FindRelationship(XDocument document, string id,
        string relationshipPart)
    {
        if (string.IsNullOrWhiteSpace(id))
            return null;
        var item = document.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "Relationship"
            && string.Equals((string)element.Attribute("Id"), id, StringComparison.Ordinal));
        if (item == null)
            return null;
        var target = (string)item.Attribute("Target");
        var external = string.Equals((string)item.Attribute("TargetMode"), "External",
            StringComparison.OrdinalIgnoreCase);
        return new RelationshipInfo(id, (string)item.Attribute("Type"), target,
            external || string.IsNullOrWhiteSpace(target) ? null
                : ExcelPivotTemplateParts.ResolveTarget(relationshipPart, target), external);
    }

    /// <summary>
    /// 向关系文档添加唯一关系。
    /// </summary>
    private static string AddRelationship(XDocument document, RelationshipInfo original,
        string targetPart, string ownerPart)
    {
        var existing = document.Root.Elements().Where(item => item.Name.LocalName == "Relationship")
            .Select(item => (string)item.Attribute("Id")).ToHashSet(StringComparer.Ordinal);
        var id = "rIdPivot1";
        var index = 1;
        while (!existing.Add(id))
            id = "rIdPivot" + (++index).ToString(System.Globalization.CultureInfo.InvariantCulture);
        var target = new Uri("http://bing-offices.local/" + ownerPart).MakeRelativeUri(
            new Uri("http://bing-offices.local/" + Normalize(targetPart))).ToString();
        document.Root.Add(new XElement(XName.Get("Relationship", RelationshipNamespace),
            new XAttribute("Id", id), new XAttribute("Type", original.Type), new XAttribute("Target", target)));
        return id;
    }

    /// <summary>
    /// 获取关系部件对应的拥有者路径。
    /// </summary>
    private static string RelationshipOwner(string relationshipPart)
    {
        var name = Normalize(relationshipPart);
        if (string.Equals(name, "_rels/.rels", StringComparison.OrdinalIgnoreCase))
            return string.Empty;
        var marker = name.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        if (marker >= 0)
            return name[..(marker + 1)] + name[(marker + 7)..^5];
        if (name.StartsWith("_rels/", StringComparison.OrdinalIgnoreCase)
            && name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return name[6..^5];
        throw new InvalidDataException($"关系部件路径无效: {relationshipPart}");
    }

    /// <summary>
    /// 计算关系目标相对于拥有者的路径。
    /// </summary>
    private static string RelativeTarget(string ownerPart, string targetPart)
    {
        var origin = new Uri("http://bing-offices.local/" + Normalize(ownerPart));
        var target = new Uri("http://bing-offices.local/" + Normalize(targetPart));
        return origin.MakeRelativeUri(target).ToString();
    }

    /// <summary>
    /// 获取关系文档根元素。
    /// </summary>
    private static XDocument EnsureRelationships(XDocument document)
    {
        if (document.Root == null || document.Root.Name.LocalName != "Relationships")
            throw new InvalidDataException("关系部件根元素无效。");
        return document;
    }

    /// <summary>
    /// 创建空关系文档。
    /// </summary>
    private static XDocument NewRelationships() => new(new XElement(XName.Get("Relationships", RelationshipNamespace)));

    /// <summary>
    /// 计算关系部件名称。
    /// </summary>
    private static string RelationshipPart(string owner)
    {
        var index = owner.LastIndexOf('/');
        return owner[..(index + 1)] + "_rels/" + owner[(index + 1)..] + ".rels";
    }

    /// <summary>
    /// 规范化包部件名称。
    /// </summary>
    private static string Normalize(string name) => name?.Replace('\\', '/').TrimStart('/') ?? string.Empty;

    /// <summary>
    /// 为同步 XML 读取提供可中断的只读流包装。
    /// </summary>
    private sealed class CancellationReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly CancellationToken _token;

        internal CancellationReadStream(Stream inner, CancellationToken token)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _token = token;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => _inner.CanSeek;
        public override bool CanWrite => _inner.CanWrite;
        public override long Length => _inner.Length;
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        public override void Flush() => _inner.Flush();

        public override int Read(byte[] buffer, int offset, int count)
        {
            _token.ThrowIfCancellationRequested();
            var read = _inner.Read(buffer, offset, count);
            _token.ThrowIfCancellationRequested();
            return read;
        }

        public override int Read(Span<byte> buffer)
        {
            _token.ThrowIfCancellationRequested();
            var read = _inner.Read(buffer);
            _token.ThrowIfCancellationRequested();
            return read;
        }

        public override int ReadByte()
        {
            _token.ThrowIfCancellationRequested();
            var value = _inner.ReadByte();
            _token.ThrowIfCancellationRequested();
            return value;
        }

        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        public override void SetLength(long value) => _inner.SetLength(value);
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 判断关系部件是否属于透视部件本身。
    /// </summary>
    private static bool IsPivotRelationshipPart(string name, ISet<string> pivotParts)
    {
        if (!name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return false;
        var marker = name.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
        var owner = marker < 0 ? null : name[..marker] + "/" + name[(marker + 7)..^5];
        return owner != null && pivotParts.Contains(owner);
    }

    private sealed class PackageInfo
    {
        internal PackageInfo(string workbookPart, string workbookRelationshipsPart,
            XDocument workbook, XDocument workbookRelationships,
            Dictionary<string, string> sheetParts, Dictionary<string, XDocument> sheetDocuments,
            Dictionary<string, XDocument> sheetRelationships)
        {
            WorkbookPart = workbookPart;
            WorkbookRelationshipsPart = workbookRelationshipsPart;
            Workbook = workbook;
            WorkbookRelationships = workbookRelationships;
            SheetParts = sheetParts;
            SheetDocuments = sheetDocuments;
            SheetRelationships = sheetRelationships;
        }

        internal string WorkbookPart { get; }
        internal string WorkbookRelationshipsPart { get; }
        internal XDocument Workbook { get; }
        internal XDocument WorkbookRelationships { get; }
        internal Dictionary<string, string> SheetParts { get; }
        internal Dictionary<string, XDocument> SheetDocuments { get; }
        internal Dictionary<string, XDocument> SheetRelationships { get; }
    }

    private sealed class RelationshipInfo
    {
        internal RelationshipInfo(string id, string type, string target, string partName, bool external)
        {
            Id = id;
            Type = type;
            Target = target;
            PartName = partName;
            External = external;
        }

        internal string Id { get; }
        internal string Type { get; }
        internal string Target { get; }
        internal string PartName { get; }
        internal bool External { get; }
    }
}
