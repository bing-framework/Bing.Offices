using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.PackageConsumer;

/// <summary>
/// 验证已打包的基础 Workbook 透视 API 保存完整来源与原生缓存。
/// </summary>
internal static class BasicPivotPackageContract
{
    /// <summary>
    /// 验证 Provider 包中的同步与异步基础透视导出。
    /// </summary>
    /// <param name="exporter">提供 Workbook 导出的公开 Provider 接口。</param>
    public static async Task VerifyAsync(IExcelExporter exporter)
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelProviderFeatureDescriptor features
            && !features.Supports(ExcelProviderFeatures.PivotTables))
            throw new InvalidOperationException("Consumer exporter does not declare PivotTables.");

        using var syncOutput = new MemoryStream();
        exporter.Export(CreateRequest(), syncOutput);
        Ensure(syncOutput.CanRead && syncOutput.CanWrite,
            "Synchronous exporter closed the caller-owned output stream.");
        AssertWorkbook(syncOutput.ToArray(), "synchronous");

        using var asyncOutput = new MemoryStream();
        await exporter.ExportAsync(CreateRequest(), asyncOutput).ConfigureAwait(false);
        Ensure(asyncOutput.CanRead && asyncOutput.CanWrite,
            "Asynchronous exporter closed the caller-owned output stream.");
        AssertWorkbook(asyncOutput.ToArray(), "asynchronous");
    }

    /// <summary>
    /// 创建使用旧式行字段和值字段配置的基础透视请求。
    /// </summary>
    /// <returns>公开 Workbook 构建入口创建的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest() =>
        ExcelExport.Workbook(builder => builder
            .Format(ExcelFormat.Xlsx)
            .AddSheet("Sales", new[]
            {
                new BasicPivotPackageRow { Warehouse = "North", Amount = 10m, Item = "A-1" },
                new BasicPivotPackageRow { Warehouse = "North", Amount = 20m, Item = "A-2" },
                new BasicPivotPackageRow { Warehouse = "South", Amount = 30m, Item = "B-1" }
            })
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "ConsumerBasicPivot",
                SourceSheet = "Sales",
                SourceRange = new ExcelRangeDefinition
                    { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 2 },
                TargetSheet = "Summary",
                RowField = "Warehouse",
                ValueField = "Amount"
            }));

    /// <summary>
    /// 检查重开的 XLSX 工作簿对象模型与完整缓存包部件。
    /// </summary>
    /// <param name="bytes">Provider 导出的 XLSX 内容。</param>
    /// <param name="operation">用于失败信息标识的导出路径。</param>
    private static void AssertWorkbook(byte[] bytes, string operation)
    {
        Ensure(bytes != null && bytes.Length > 0, $"{operation}: exporter produced an empty workbook.");
        using (var workbook = new XSSFWorkbook(new MemoryStream(bytes)))
        {
            Ensure(Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName)
                    .SequenceEqual(new[] { "Sales", "Summary" }),
                $"{operation}: workbook sheet order changed.");
            var source = workbook.GetSheet("Sales");
            var expected = ExpectedRows();
            Ensure(source.GetRow(0).Cells.Take(3).Select(cell => cell.StringCellValue)
                    .SequenceEqual(new[] { "Warehouse", "Amount", "Item" }),
                $"{operation}: source headers changed.");
            for (var index = 0; index < expected.Length; index++)
            {
                var row = source.GetRow(index + 1);
                Ensure(row.GetCell(0).StringCellValue == expected[index][0]
                    && row.GetCell(1).NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        == expected[index][1]
                    && row.GetCell(2).StringCellValue == expected[index][2],
                    $"{operation}: source row {index + 1} changed.");
            }

            var pivot = ((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables().SingleOrDefault();
            Ensure(pivot != null && pivot.GetCTPivotTableDefinition().name == "ConsumerBasicPivot",
                $"{operation}: native pivot table did not reopen with its configured name.");
        }

        AssertPackage(bytes, operation);
    }

    /// <summary>
    /// 验证工作簿、工作表、透视表与完整缓存记录的 OPC 关系。
    /// </summary>
    /// <param name="bytes">Provider 导出的 XLSX 内容。</param>
    /// <param name="operation">用于失败信息标识的导出路径。</param>
    private static void AssertPackage(byte[] bytes, string operation)
    {
        using var input = new MemoryStream(bytes);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        var parts = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var stream = entry.Open();
                return XDocument.Load(stream);
            }, StringComparer.Ordinal);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relationshipId = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        var rootRelationships = parts["_rels/.rels"];
        var workbookPath = ResolveTarget(string.Empty, rootRelationships,
            RelationshipId(rootRelationships, "officeDocument"), parts);
        Ensure(workbookPath == "xl/workbook.xml", $"{operation}: workbook relationship target changed.");

        var workbook = parts[workbookPath];
        var workbookRelationships = parts[RelationshipPartPath(workbookPath)];
        var sheetElements = workbook.Root.Element(ns + "sheets").Elements(ns + "sheet").ToArray();
        Ensure(sheetElements.Select(sheet => (string)sheet.Attribute("name"))
                .SequenceEqual(new[] { "Sales", "Summary" }),
            $"{operation}: workbook sheet names changed.");
        var sheetPaths = sheetElements.ToDictionary(sheet => (string)sheet.Attribute("name"),
            sheet => ResolveTarget(workbookPath, workbookRelationships,
                (string)sheet.Attribute(relationshipId + "id"), parts), StringComparer.Ordinal);

        var summaryRelationships = parts[RelationshipPartPath(sheetPaths["Summary"])];
        var pivotTableId = RelationshipId(summaryRelationships, "pivotTable");
        var pivotTablePath = ResolveTarget(sheetPaths["Summary"], summaryRelationships, pivotTableId, parts);
        var pivot = parts[pivotTablePath].Root;
        Ensure((string)pivot.Attribute("name") == "ConsumerBasicPivot",
            $"{operation}: pivot table name changed.");
        Ensure(pivot.Element(ns + "rowFields").Elements(ns + "field")
                .Select(field => (int)field.Attribute("x")).SequenceEqual(new[] { 0 }),
            $"{operation}: legacy row-field index changed.");
        var valueField = pivot.Element(ns + "dataFields").Elements(ns + "dataField").Single();
        Ensure((int)valueField.Attribute("fld") == 1
            && ((string)valueField.Attribute("subtotal") ?? "sum") == "sum",
            $"{operation}: legacy value field or sum aggregation changed.");

        var cacheId = RelationshipId(workbookRelationships, "pivotCacheDefinition");
        var cachePath = ResolveTarget(workbookPath, workbookRelationships, cacheId, parts);
        var pivotRelationships = parts[RelationshipPartPath(pivotTablePath)];
        var pivotCachePath = ResolveTarget(pivotTablePath, pivotRelationships,
            RelationshipId(pivotRelationships, "pivotCacheDefinition"), parts);
        Ensure(pivotCachePath == cachePath,
            $"{operation}: pivot-to-cache relationship changed.");
        var cache = parts[cachePath].Root;
        Ensure((int)cache.Attribute("recordCount") == 3,
            $"{operation}: cache record count changed.");
        Ensure((string)cache.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("sheet")
                == "Sales"
            && (string)cache.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref")
                == "A1:C4",
            $"{operation}: cache source sheet or range changed.");
        var cacheFields = cache.Element(ns + "cacheFields").Elements(ns + "cacheField").ToArray();
        Ensure(cacheFields.Select(field => (string)field.Attribute("name"))
                .SequenceEqual(new[] { "Warehouse", "Amount", "Item" }),
            $"{operation}: cache field names or order changed.");

        var cacheRelationships = parts[RelationshipPartPath(cachePath)];
        var recordsPath = ResolveTarget(cachePath, cacheRelationships,
            RelationshipId(cacheRelationships, "pivotCacheRecords"), parts);
        var records = parts[recordsPath].Root.Elements(ns + "r").ToArray();
        var expected = ExpectedRows();
        Ensure((int)parts[recordsPath].Root.Attribute("count") == expected.Length
            && records.Length == expected.Length,
            $"{operation}: cache record count does not match all source rows.");
        for (var rowIndex = 0; rowIndex < expected.Length; rowIndex++)
        {
            var cells = records[rowIndex].Elements().ToArray();
            Ensure(cells.Length == 3, $"{operation}: cache row {rowIndex + 1} field count changed.");
            var values = cells.Select((cell, column) => ReadRecordValue(cell, cacheFields[column], ns)).ToArray();
            Ensure(values.SequenceEqual(expected[rowIndex]),
                $"{operation}: cache row {rowIndex + 1} values changed.");
        }

        var contentTypes = parts["[Content_Types].xml"].Root;
        foreach (var path in new[] { pivotTablePath, cachePath, recordsPath })
            Ensure(contentTypes.Elements().Any(element =>
                    (string)element.Attribute("PartName") == "/" + path),
                $"{operation}: content type override missing for {path}.");
    }

    /// <summary>
    /// 解析来源行的缓存值或共享项目索引。
    /// </summary>
    /// <param name="cell">缓存记录中的单元格。</param>
    /// <param name="field">对应的缓存字段。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <returns>缓存中的逻辑值。</returns>
    private static string ReadRecordValue(XElement cell, XElement field, XNamespace ns)
    {
        var value = (string)cell.Attribute("v");
        if (cell.Name == ns + "x")
            return (string)field.Element(ns + "sharedItems").Elements()
                .ElementAt(int.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).Attribute("v");
        return value;
    }

    /// <summary>
    /// 解析关系目标，并验证其对应部件存在。
    /// </summary>
    /// <param name="sourcePart">关系所属部件。</param>
    /// <param name="relationships">关系 XML。</param>
    /// <param name="id">关系标识。</param>
    /// <param name="parts">工作簿中已读取的 XML 部件。</param>
    /// <returns>规范化后的目标部件路径。</returns>
    private static string ResolveTarget(string sourcePart, XDocument relationships, string id,
        IReadOnlyDictionary<string, XDocument> parts)
    {
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var target = (string)relationships.Root.Elements(relNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == id).Attribute("Target");
        var segments = target.StartsWith("/", StringComparison.Ordinal)
            ? new List<string>()
            : sourcePart.Split('/').SkipLast(1).ToList();
        foreach (var segment in target.TrimStart('/').Split('/'))
        {
            if (segment == "..")
                segments.RemoveAt(segments.Count - 1);
            else if (segment != "." && segment.Length > 0)
                segments.Add(segment);
        }
        var path = string.Join("/", segments);
        Ensure(parts.ContainsKey(path), $"Workbook relationship target is missing: {path}.");
        return path;
    }

    /// <summary>
    /// 根据关系类型后缀读取唯一的关系标识。
    /// </summary>
    /// <param name="relationships">关系 XML。</param>
    /// <param name="typeSuffix">关系类型 URI 后缀。</param>
    /// <returns>匹配关系的标识。</returns>
    private static string RelationshipId(XDocument relationships, string typeSuffix)
    {
        XNamespace relNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        return (string)relationships.Root.Elements(relNs + "Relationship")
            .Single(item => ((string)item.Attribute("Type")).EndsWith(typeSuffix, StringComparison.Ordinal))
            .Attribute("Id");
    }

    /// <summary>
    /// 根据部件路径生成其 OPC 关系部件路径。
    /// </summary>
    /// <param name="partPath">来源部件路径。</param>
    /// <returns>对应关系部件路径。</returns>
    private static string RelationshipPartPath(string partPath)
    {
        var slash = partPath.LastIndexOf('/');
        return partPath.Substring(0, slash + 1) + "_rels/" + partPath.Substring(slash + 1) + ".rels";
    }

    /// <summary>
    /// 创建基础透视合同使用的完整来源记录。
    /// </summary>
    /// <returns>包含全部仓库、金额和商品编号的预期来源值。</returns>
    private static string[][] ExpectedRows() => new[]
    {
        new[] { "North", "10", "A-1" },
        new[] { "North", "20", "A-2" },
        new[] { "South", "30", "B-1" }
    };

    /// <summary>
    /// 在合同失败时抛出带有原因的异常。
    /// </summary>
    /// <param name="condition">必须满足的合同条件。</param>
    /// <param name="message">条件不满足时的错误说明。</param>
    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
