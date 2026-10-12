using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using ClosedXML.Excel;

namespace Bing.Offices.PackageConsumer;

/// <summary>
/// 验证已打包的公开 Workbook 透视 API 保存完整配置与缓存结构。
/// </summary>
internal static class AdvancedPivotPackageContract
{
    /// <summary>
    /// 获取来源表、值字段和过滤后仍保留在缓存中的全部记录。
    /// </summary>
    private static readonly string[] ExpectedRecords =
    {
        "Warehouse", "Category", "Month", "Status", "Amount", "Item",
        "North", "Food", "Jan", "Open", "10", "A",
        "North", "Food", "Feb", "Closed", "30", "B",
        "North", "Tools", "Jan", "Open", "20", "C",
        "South", "Food", "Jan", "Pending", "40", "D"
    };

    /// <summary>
    /// 验证已发布包中的同步与异步高级透视导出合同。
    /// </summary>
    /// <param name="exporter">提供 Workbook 导出的公开 Provider 接口。</param>
    /// <returns>包含同步和异步结构验证的任务。</returns>
    public static async Task VerifyAsync(IExcelExporter exporter)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelProviderFeatureDescriptor features
            && !features.Supports(ExcelProviderFeatures.AdvancedPivotTables))
            throw new InvalidOperationException("Consumer exporter does not declare AdvancedPivotTables.");

        using var syncOutput = new MemoryStream();
        exporter.Export(CreateRequest(), syncOutput);
        Ensure(syncOutput.CanWrite, "Synchronous exporter closed the caller-owned output stream.");
        AssertWorkbook(syncOutput.ToArray(), "synchronous");

        using var asyncOutput = new MemoryStream();
        await exporter.ExportAsync(CreateRequest(), asyncOutput).ConfigureAwait(false);
        Ensure(asyncOutput.CanWrite, "Asynchronous exporter closed the caller-owned output stream.");
        AssertWorkbook(asyncOutput.ToArray(), "asynchronous");
    }

    /// <summary>
    /// 验证已发布包中的透视模板更新入口。
    /// </summary>
    /// <param name="exporter">提供 Workbook 导出的公开 Provider 接口。</param>
    /// <returns>包含同步和异步更新检查的任务。</returns>
    public static async Task VerifyTemplateUpdateAsync(IExcelExporter exporter)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelProviderFeatureDescriptor features
            && !features.Supports(ExcelProviderFeatures.TemplatePivotTables))
            throw new InvalidOperationException("Consumer exporter does not declare TemplatePivotTables.");

        await VerifyTemplateUpdateAsync(exporter, asynchronous: false).ConfigureAwait(false);
        await VerifyTemplateUpdateAsync(exporter, asynchronous: true).ConfigureAwait(false);
    }

    /// <summary>
    /// 执行单一同步或异步模板更新检查。
    /// </summary>
    /// <param name="exporter">提供 Workbook 导出的公开 Provider 接口。</param>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    private static async Task VerifyTemplateUpdateAsync(IExcelExporter exporter, bool asynchronous)
    {
        using var templateOutput = new MemoryStream();
        exporter.Export(CreateTemplateRequest(), templateOutput);
        var templateBytes = templateOutput.ToArray();

        using var template = new MemoryStream(templateBytes, writable: false);
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .UpdatePivotTable(CreateTemplatePivotDefinition()));
        using var output = new MemoryStream();

        if (asynchronous)
            await exporter.ExportAsync(request, output).ConfigureAwait(false);
        else
            exporter.Export(request, output);

        Ensure(template.CanRead && template.ToArray().SequenceEqual(templateBytes),
            "Template pivot update changed or closed the caller-owned template stream.");
        AssertTemplateUpdate(output.ToArray(), asynchronous ? "asynchronous" : "synchronous");
    }

    /// <summary>
    /// 创建包含基础透视表的消费者模板。
    /// </summary>
    /// <returns>可供公开模板更新入口使用的导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateTemplateRequest() =>
        ExcelExport.Workbook(builder => builder
            .Format(ExcelFormat.Xlsx)
            .AddSheet("TemplateData", new[]
            {
                new BasicPivotPackageRow { Warehouse = "North", Amount = 10, Item = "A-1" },
                new BasicPivotPackageRow { Warehouse = "South", Amount = 20, Item = "B-1" },
                new BasicPivotPackageRow { Warehouse = "North", Amount = 30, Item = "A-2" }
            })
            .PivotTable(CreateTemplatePivotDefinition()));

    /// <summary>
    /// 创建模板透视表的更新定义。
    /// </summary>
    /// <returns>使用模板中来源和目标坐标的透视定义。</returns>
    private static ExcelPivotTableDefinition CreateTemplatePivotDefinition() => new()
    {
        Name = "PackageTemplatePivot",
        SourceSheet = "TemplateData",
        SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 2 },
        TargetSheet = "TemplateSummary",
        RowField = "Warehouse",
        ValueField = "Amount"
    };

    /// <summary>
    /// 检查模板更新后的透视定义、缓存刷新标记和完整记录。
    /// </summary>
    /// <param name="bytes">模板更新导出的 XLSX 字节。</param>
    /// <param name="operation">当前同步或异步入口名称。</param>
    private static void AssertTemplateUpdate(byte[] bytes, string operation)
    {
        Ensure(bytes != null && bytes.Length > 0, $"{operation}: template update produced an empty workbook.");
        using var input = new MemoryStream(bytes);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var pivot = ReadXml(archive, entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && entry.FullName.Contains("pivotTables/", StringComparison.OrdinalIgnoreCase));
        Ensure((string)pivot.Root.Attribute("name") == "PackageTemplatePivot",
            $"{operation}: updated template pivot name changed.");

        var cache = ReadXml(archive, entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && entry.FullName.Contains("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase));
        Ensure(XmlConvert.ToBoolean((string)cache.Root.Attribute("refreshOnLoad") ?? "false"),
            $"{operation}: updated cache did not request refresh on load.");
        var cacheFields = cache.Root.Element(ns + "cacheFields").Elements(ns + "cacheField").ToArray();
        Ensure(cacheFields.Select(field => (string)field.Attribute("name"))
                .SequenceEqual(new[] { "Warehouse", "Amount", "Item" }),
            $"{operation}: updated cache fields changed.");

        var records = ReadXml(archive, entry => entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
            && entry.FullName.Contains("pivotCacheRecords", StringComparison.OrdinalIgnoreCase));
        var rows = records.Root.Elements(ns + "r").ToArray();
        var declaredCount = (string)records.Root.Attribute("count");
        Ensure(rows.Length == 3 && (declaredCount == null
                || XmlConvert.ToInt32(declaredCount) == rows.Length),
            $"{operation}: updated cache row count changed.");
        var values = rows.SelectMany(row => row.Elements()
            .Select((cell, index) => ReadCacheValue(cell, cacheFields[index], ns))).ToArray();
        Ensure(values.SequenceEqual(new[]
        {
            "North", "10", "A-1",
            "South", "20", "B-1",
            "North", "30", "A-2"
        }), $"{operation}: updated cache values changed.");
    }

    /// <summary>
    /// 从工作簿包中读取唯一符合条件的 XML 部件。
    /// </summary>
    /// <param name="archive">待检查的 XLSX 包。</param>
    /// <param name="predicate">部件选择条件。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, Func<ZipArchiveEntry, bool> predicate)
    {
        var entry = archive.Entries.SingleOrDefault(predicate);
        Ensure(entry != null, "Template update output is missing a required pivot package part.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 创建包含多轴、五种聚合、筛选及不同数字格式的透视请求。
    /// </summary>
    /// <returns>公开 Workbook 构建入口创建的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest() =>
        ExcelExport.Workbook(builder => builder
            .Format(ExcelFormat.Xlsx)
            .AddSheet("Data", new[]
            {
                new AdvancedPivotPackageRow { Warehouse = "North", Category = "Food", Month = "Jan", Status = "Open", Amount = 10, Item = "A" },
                new AdvancedPivotPackageRow { Warehouse = "North", Category = "Food", Month = "Feb", Status = "Closed", Amount = 30, Item = "B" },
                new AdvancedPivotPackageRow { Warehouse = "North", Category = "Tools", Month = "Jan", Status = "Open", Amount = 20, Item = "C" },
                new AdvancedPivotPackageRow { Warehouse = "South", Category = "Food", Month = "Jan", Status = "Pending", Amount = 40, Item = "D" }
            })
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "PackageAdvancedPivot",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 4, EndColumn = 5 },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Warehouse", SortOrder = ExcelPivotSortOrder.Ascending },
                    new ExcelPivotAxisFieldDefinition { Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                        SubtotalMode = ExcelPivotSubtotalMode.None }
                },
                ColumnFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Month", SortOrder = ExcelPivotSortOrder.Ascending,
                        SubtotalMode = ExcelPivotSubtotalMode.None }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", Aggregation = ExcelPivotAggregation.Sum, NumberFormat = "0.00" },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Mean", Aggregation = ExcelPivotAggregation.Average, NumberFormat = "0.0" },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Low", Aggregation = ExcelPivotAggregation.Minimum, NumberFormat = "0.000\"kg\"" },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "High", Aggregation = ExcelPivotAggregation.Maximum },
                    new ExcelPivotValueDefinition { Field = "Item", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
                },
                Filters = new[]
                {
                    new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { "Open", "Pending" } }
                }
            }));

    /// <summary>
    /// 检查重开的 XLSX 工作簿对象模型及 OOXML 部件。
    /// </summary>
    /// <param name="bytes">导出的 XLSX 字节。</param>
    /// <param name="operation">用于失败信息标识的导出路径。</param>
    private static void AssertWorkbook(byte[] bytes, string operation)
    {
        Ensure(bytes != null && bytes.Length > 0, $"{operation}: exporter produced an empty workbook.");
        using (var workbook = new XLWorkbook(new MemoryStream(bytes)))
        {
            Ensure(workbook.Worksheets.Select(sheet => sheet.Name)
                    .SequenceEqual(new[] { "Data", "Summary" }),
                $"{operation}: workbook sheet order changed.");
            Ensure(workbook.Worksheet("Data").Range("A1:F5").Cells().Select(cell => cell.GetString())
                    .SequenceEqual(ExpectedRecords),
                $"{operation}: source worksheet values changed.");

            var pivot = workbook.Worksheet("Summary").PivotTables.SingleOrDefault();
            Ensure(pivot != null && pivot.Name == "PackageAdvancedPivot",
                $"{operation}: expected pivot table is missing.");
            Ensure(pivot.RowLabels.Select(field => field.SourceName)
                    .SequenceEqual(new[] { "Warehouse", "Category" }),
                $"{operation}: row field order changed.");
            Ensure(pivot.ColumnLabels.Any(field => field.SourceName == "Month"),
                $"{operation}: month column field is missing.");
            Ensure(pivot.RowLabels.Get("Warehouse").SortType == XLPivotSortType.Ascending
                && pivot.RowLabels.Get("Category").SortType == XLPivotSortType.Descending
                && pivot.ColumnLabels.Get("Month").SortType == XLPivotSortType.Ascending,
                $"{operation}: axis sorting was not preserved.");
            Ensure(pivot.RowLabels.Get("Warehouse").Subtotals
                    .SequenceEqual(new[] { XLSubtotalFunction.Automatic })
                && !pivot.RowLabels.Get("Category").Subtotals.Any()
                && !pivot.ColumnLabels.Get("Month").Subtotals.Any(),
                $"{operation}: axis subtotal configuration changed.");
            Ensure(pivot.Values.Select(value => value.CustomName)
                    .SequenceEqual(new[] { "Total", "Mean", "Low", "High", "Count" }),
                $"{operation}: value field order or captions changed.");
            Ensure(pivot.Values.Select(value => value.SummaryFormula)
                    .SequenceEqual(new[] { XLPivotSummary.Sum, XLPivotSummary.Average,
                        XLPivotSummary.Minimum, XLPivotSummary.Maximum, XLPivotSummary.Count }),
                $"{operation}: value aggregation changed.");
            Ensure(pivot.Values.Select(GetNumberFormat)
                    .SequenceEqual(new[] { "0.00", "0.0", "0.000\"kg\"", null, null }),
                $"{operation}: value number formats were not isolated: "
                + string.Join(" | ", pivot.Values.Select(value =>
                    $"{value.CustomName}={value.NumberFormat.NumberFormatId}:{value.NumberFormat.Format}")));
            Ensure(pivot.Values.Select(value => value.NumberFormat.NumberFormatId)
                    .SequenceEqual(new[] { 2, -1, -1, -1, -1 }),
                $"{operation}: built-in, custom or default format identifiers changed.");
            Ensure(pivot.ReportFilters.Count() == 1
                && pivot.ReportFilters.Single().SourceName == "Status",
                $"{operation}: report filter field changed.");
        }

        AssertPackageParts(bytes, operation);
    }

    /// <summary>
    /// 检查透视定义、筛选、数字格式和未筛减的完整来源缓存。
    /// </summary>
    /// <param name="bytes">导出的 XLSX 字节。</param>
    /// <param name="operation">用于失败信息标识的导出路径。</param>
    private static void AssertPackageParts(byte[] bytes, string operation)
    {
        using var input = new MemoryStream(bytes);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        var parts = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var partStream = entry.Open();
                return XDocument.Load(partStream);
            }, StringComparer.Ordinal);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var workbook = parts["xl/workbook.xml"];
        var workbookRelationships = parts["xl/_rels/workbook.xml.rels"];

        var sheetParts = workbook.Root.Element(ns + "sheets").Elements(ns + "sheet")
            .ToDictionary(sheet => (string)sheet.Attribute("name"), sheet => ResolveRelationship(
                "xl/workbook.xml", workbookRelationships, (string)sheet.Attribute(officeRel + "id"), parts),
                StringComparer.Ordinal);
        var sourcePart = parts[sheetParts["Data"]];
        var sharedStrings = parts.TryGetValue("xl/sharedStrings.xml", out var sharedStringsPart)
            ? sharedStringsPart.Root.Elements(ns + "si")
                .Select(item => string.Concat(item.Descendants(ns + "t").Select(text => text.Value))).ToArray()
            : Array.Empty<string>();
        Ensure(sourcePart.Root.Element(ns + "sheetData").Elements(ns + "row")
                .SelectMany(row => row.Elements(ns + "c")
                    .Select(cell => ReadCellValue(cell, ns, sharedStrings))).SequenceEqual(ExpectedRecords),
            $"{operation}: worksheet XML source cells are incomplete.");

        var summaryPart = sheetParts["Summary"];
        var summaryRelationships = parts[RelationshipPartPath(summaryPart)];
        var pivotPart = ResolveRelationship(summaryPart, summaryRelationships,
            RelationshipId(summaryRelationships, "pivotTable"), parts);
        var pivot = parts[pivotPart].Root;
        Ensure((string)pivot.Attribute("name") == "PackageAdvancedPivot",
            $"{operation}: pivot table XML name changed.");
        Ensure(((string)pivot.Element(ns + "location").Attribute("ref")).Split(':')[0] == "C4"
            && (string)pivot.Element(ns + "location").Attribute("rowPageCount") == "1",
            $"{operation}: pivot body did not follow its configured filter area.");
        var fields = pivot.Element(ns + "pivotFields").Elements(ns + "pivotField").ToArray();
        Ensure(fields.Length == 6, $"{operation}: pivot field count changed.");
        var rowFields = pivot.Element(ns + "rowFields").Elements(ns + "field")
            .Select(field => (string)field.Attribute("x")).ToArray();
        var columnFields = pivot.Element(ns + "colFields").Elements(ns + "field")
            .Select(field => (string)field.Attribute("x")).ToArray();
        Ensure(rowFields.SequenceEqual(new[] { "0", "1" }), $"{operation}: row axis indices changed.");
        Ensure(columnFields.SequenceEqual(new[] { "2", "-2" }),
            $"{operation}: month and synthetic values column axis order changed.");
        var pageFields = pivot.Element(ns + "pageFields").Elements(ns + "pageField").ToArray();
        Ensure(pageFields.Length == 1 && (string)pageFields[0].Attribute("fld") == "3",
            $"{operation}: status filter field index changed.");
        var cacheId = (string)pivot.Attribute("cacheId");
        var cacheDeclaration = workbook.Root.Element(ns + "pivotCaches").Elements(ns + "pivotCache")
            .SingleOrDefault(item => (string)item.Attribute("cacheId") == cacheId);
        Ensure(cacheDeclaration != null, $"{operation}: workbook cache declaration is missing.");
        var cachePart = ResolveRelationship("xl/workbook.xml", workbookRelationships,
            (string)cacheDeclaration.Attribute(officeRel + "id"), parts);
        var cache = parts[cachePart].Root;
        var source = cache.Element(ns + "cacheSource").Element(ns + "worksheetSource");
        Ensure((string)source.Attribute("sheet") == "Data" && (string)source.Attribute("ref") == "A1:F5",
            $"{operation}: cache source range changed.");
        var cacheFields = cache.Element(ns + "cacheFields").Elements(ns + "cacheField").ToArray();
        Ensure(cacheFields.Select(field => (string)field.Attribute("name")).SequenceEqual(new[]
            { "Warehouse", "Category", "Month", "Status", "Amount", "Item" }),
            $"{operation}: complete cache field sequence changed.");
        Ensure(ReadSelectedFilterValues(fields[3], cacheFields[3], ns)
                .OrderBy(value => value, StringComparer.Ordinal)
                .SequenceEqual(new[] { "Open", "Pending" }),
            $"{operation}: selected status values changed.");

        var dataFields = pivot.Element(ns + "dataFields").Elements(ns + "dataField").ToArray();
        Ensure(dataFields.Select(field => (string)field.Attribute("name")).SequenceEqual(
                new[] { "Total", "Mean", "Low", "High", "Count" }),
            $"{operation}: OOXML value field captions changed.");
        Ensure(dataFields.Select(field => (string)field.Attribute("fld")).SequenceEqual(
                new[] { "4", "4", "4", "4", "5" }),
            $"{operation}: source field indices for aggregations changed.");
        Ensure(dataFields.Select(field => (string)field.Attribute("subtotal") ?? "sum").SequenceEqual(
                new[] { "sum", "average", "min", "max", "count" }),
            $"{operation}: OOXML aggregation types changed.");

        var styles = parts["xl/styles.xml"].Root;
        var numberFormats = styles.Element(ns + "numFmts")?.Elements(ns + "numFmt")
            .ToDictionary(item => (string)item.Attribute("numFmtId"), item => (string)item.Attribute("formatCode"))
            ?? new Dictionary<string, string>();
        var ids = dataFields.Select(field => (string)field.Attribute("numFmtId")).ToArray();
        Ensure(ids[0] == "2" && ResolveFormat(ids[0], numberFormats) == "0.00",
            $"{operation}: built-in 0.00 format mapping changed.");
        Ensure(ResolveFormat(ids[1], numberFormats) == "0.0"
            && ResolveFormat(ids[2], numberFormats) == "0.000\"kg\"",
            $"{operation}: custom value format mapping changed.");
        Ensure(ids[0] != ids[1] && ids[1] != ids[2] && ids[0] != ids[2],
            $"{operation}: value fields shared a number format identifier.");
        Ensure(ids.Skip(3).All(id => id == null || id == "0"),
            $"{operation}: unformatted fields gained an explicit number format.");

        var cacheRelationships = parts[RelationshipPartPath(cachePart)];
        var recordsPart = ResolveRelationship(cachePart, cacheRelationships,
            RelationshipId(cacheRelationships, "pivotCacheRecords"), parts);
        var records = parts[recordsPart].Root.Elements(ns + "r").ToArray();
        Ensure(records.Length == 4 && records.All(record => record.Elements().Count() == 6),
            $"{operation}: cache did not retain all filtered and unfiltered source rows.");
        var values = records.SelectMany(record => record.Elements().Select((cell, index) =>
            ReadCacheValue(cell, cacheFields[index], ns))).ToArray();
        Ensure(values.SequenceEqual(new[]
        {
            "North", "Food", "Jan", "Open", "10", "A",
            "North", "Food", "Feb", "Closed", "30", "B",
            "North", "Tools", "Jan", "Open", "20", "C",
            "South", "Food", "Jan", "Pending", "40", "D"
        }), $"{operation}: cache records or values changed.");
    }

    /// <summary>
    /// 解析关系目标并要求目标部件存在于工作簿包中。
    /// </summary>
    /// <param name="sourcePart">关系所属源部件。</param>
    /// <param name="relationships">源部件关系 XML。</param>
    /// <param name="relationshipId">关系标识。</param>
    /// <param name="parts">工作簿中的 XML 部件。</param>
    /// <returns>规范化后的关系目标路径。</returns>
    private static string ResolveRelationship(string sourcePart, XDocument relationships,
        string relationshipId, IReadOnlyDictionary<string, XDocument> parts)
    {
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var target = (string)relationships.Root.Elements(relNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == relationshipId).Attribute("Target");
        var path = target.StartsWith("/", StringComparison.Ordinal)
            ? target.TrimStart('/')
            : string.Join("/", sourcePart.Split('/').SkipLast(1).Concat(target.Split('/')));
        var segments = new List<string>();
        foreach (var segment in path.Split('/'))
        {
            if (segment == "..")
                segments.RemoveAt(segments.Count - 1);
            else if (segment != "." && segment.Length != 0)
                segments.Add(segment);
        }
        var normalized = string.Join("/", segments);
        Ensure(parts.ContainsKey(normalized), $"Workbook relationship target is missing: {normalized}.");
        return normalized;
    }

    /// <summary>
    /// 按目标关系类型读取唯一的关系标识。
    /// </summary>
    /// <param name="relationships">关系 XML。</param>
    /// <param name="typeSuffix">关系类型 URI 后缀。</param>
    /// <returns>匹配关系的 ID。</returns>
    private static string RelationshipId(XDocument relationships, string typeSuffix)
    {
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        return (string)relationships.Root.Elements(relNs + "Relationship")
            .Single(item => ((string)item.Attribute("Type")).EndsWith(typeSuffix, StringComparison.Ordinal))
            .Attribute("Id");
    }

    /// <summary>
    /// 根据部件路径生成其关系文件路径。
    /// </summary>
    /// <param name="partPath">部件路径。</param>
    /// <returns>对应的关系部件路径。</returns>
    private static string RelationshipPartPath(string partPath)
    {
        var slash = partPath.LastIndexOf('/');
        return partPath.Substring(0, slash + 1) + "_rels/" + partPath.Substring(slash + 1) + ".rels";
    }

    /// <summary>
    /// 读取页筛选项并解析它们指向的缓存共享值。
    /// </summary>
    /// <param name="pivotField">透视字段 XML。</param>
    /// <param name="cacheField">对应缓存字段 XML。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <returns>当前选择的筛选值。</returns>
    private static string[] ReadSelectedFilterValues(XElement pivotField, XElement cacheField, XNamespace ns)
    {
        var cacheItems = cacheField.Element(ns + "sharedItems").Elements().ToArray();
        return pivotField.Element(ns + "items").Elements(ns + "item")
            .Where(item => item.Attribute("x") != null
                && !IsTrueBoolean((string)item.Attribute("h")))
            .Select(item => cacheItems[int.Parse((string)item.Attribute("x"))])
            .Select(item => (string)item.Attribute("v")).ToArray();
    }

    /// <summary>
    /// 解析 OOXML 布尔属性并判断其是否表示真。
    /// </summary>
    /// <param name="value">属性文本；缺失值按假处理。</param>
    /// <returns>属性值为 1 或 true 时返回 true。</returns>
    private static bool IsTrueBoolean(string value) =>
        value != null && XmlConvert.ToBoolean(value);

    /// <summary>
    /// 读取工作表 XML 单元格的共享字符串或原始值。
    /// </summary>
    /// <param name="cell">单元格 XML 元素。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <param name="sharedStrings">工作簿共享字符串表。</param>
    /// <returns>单元格逻辑值。</returns>
    private static string ReadCellValue(XElement cell, XNamespace ns, IReadOnlyList<string> sharedStrings)
    {
        var type = (string)cell.Attribute("t");
        if (type == "inlineStr")
            return string.Concat(cell.Descendants(ns + "t").Select(item => item.Value));
        var value = (string)cell.Element(ns + "v");
        return type == "s" ? sharedStrings[int.Parse(value)] : value;
    }

    /// <summary>
    /// 读取缓存记录值，并解析缓存字段的共享项目索引。
    /// </summary>
    /// <param name="cell">缓存记录值元素。</param>
    /// <param name="field">对应缓存字段。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <returns>缓存记录中的逻辑值。</returns>
    private static string ReadCacheValue(XElement cell, XElement field, XNamespace ns)
    {
        var value = (string)cell.Attribute("v");
        if (cell.Name == ns + "x")
            return (string)field.Element(ns + "sharedItems").Elements().ElementAt(int.Parse(value)).Attribute("v");
        return value;
    }

    /// <summary>
    /// 将透视值字段格式编号还原为内建或自定义格式文本。
    /// </summary>
    /// <param name="formatId">OOXML 数字格式编号。</param>
    /// <param name="customFormats">自定义格式编号映射。</param>
    /// <returns>对应格式文本；不支持的编号返回 <see langword="null"/>。</returns>
    private static string ResolveFormat(string formatId, IReadOnlyDictionary<string, string> customFormats)
    {
        if (formatId != null && customFormats.TryGetValue(formatId, out var format))
            return format;
        return formatId == "2" ? "0.00" : null;
    }

    /// <summary>
    /// 读取透视字段格式，按 OOXML 内建编号解析格式 0.00。
    /// </summary>
    /// <param name="value">透视值字段。</param>
    /// <returns>内建格式或自定义格式文本；默认格式返回 <see langword="null"/>。</returns>
    private static string GetNumberFormat(IXLPivotValue value)
    {
        if (value.NumberFormat.NumberFormatId == 2)
            return "0.00";
        if (value.NumberFormat.NumberFormatId == -1 && value.NumberFormat.Format.Length == 0)
            return null;
        return value.NumberFormat.Format;
    }

    /// <summary>
    /// 在合同失败时抛出带有明确原因的异常。
    /// </summary>
    /// <param name="condition">必须为真的合同条件。</param>
    /// <param name="message">条件不满足时的错误说明。</param>
    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
