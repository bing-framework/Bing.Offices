using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证高级透视中同一来源字段的多个数字格式彼此独立并完整保存。
/// </summary>
public sealed class AdvancedPivotNumberFormatTest
{
    /// <summary>
    /// 获取测试中依次用于内建及自定义度量的数字格式。
    /// </summary>
    private static readonly string[] ExpectedFormats = { "0.00", "0.0", "0.000\"kg\"" };

    /// <summary>
    /// 获取与数字格式顺序对应的透视度量标题。
    /// </summary>
    private static readonly string[] ExpectedCaptions = { "BuiltIn", "OneDecimal", "Kilograms" };

    /// <summary>
    /// 验证同一金额来源的内建和自定义格式在完整工作簿中保持隔离。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RepeatedValueFieldFormats_ShouldPreserveCompleteWorkbook(bool asynchronous)
    {
        var request = CreateRequest();
        using var output = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();
        if (asynchronous)
            await exporter.ExportAsync(request, output);
        else
            exporter.Export(request, output);

        Assert.True(output.CanWrite);
        var bytes = output.ToArray();
        using (var workbook = new XLWorkbook(new MemoryStream(bytes)))
        {
            Assert.Equal(new[] { "Data", "Summary" }, workbook.Worksheets.Select(sheet => sheet.Name));
            Assert.Equal(new[]
            {
                "Category", "Amount", "OrderId",
                "North", "1.25", "PO-1",
                "South", "2.5", "PO-2",
                "North", "3.75", "PO-3"
            }, workbook.Worksheet("Data").Range("A1:C4").Cells().Select(cell => cell.GetString()));

            var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
            Assert.Equal("FormatSummary", pivot.Name);
            Assert.Equal("Category", Assert.Single(pivot.RowLabels).SourceName);
            Assert.Equal(ExpectedCaptions, pivot.Values.Select(value => value.CustomName));
            Assert.Equal(new[] { XLPivotSummary.Sum, XLPivotSummary.Sum, XLPivotSummary.Sum },
                pivot.Values.Select(value => value.SummaryFormula));
            Assert.Equal(new[] { 2, -1, -1 }, pivot.Values.Select(value => value.NumberFormat.NumberFormatId));
            Assert.Equal(ExpectedFormats, pivot.Values.Select(GetNumberFormat));
        }

        AssertPackageFormatsAndCache(bytes);
    }

    /// <summary>
    /// 创建包含三个同源金额度量的透视导出请求。
    /// </summary>
    /// <returns>包含完整来源数据与独立格式定义的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest() =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new AdvancedPivotNumberFormatRow { Category = "North", Amount = 1.25, OrderId = "PO-1" },
            new AdvancedPivotNumberFormatRow { Category = "South", Amount = 2.5, OrderId = "PO-2" },
            new AdvancedPivotNumberFormatRow { Category = "North", Amount = 3.75, OrderId = "PO-3" }
        }).PivotTable(new ExcelPivotTableDefinition
        {
            Name = "FormatSummary",
            SourceSheet = "Data",
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 2 },
            TargetSheet = "Summary",
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Category" } },
            ValueFields = new[]
            {
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "BuiltIn", NumberFormat = "0.00" },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "OneDecimal", NumberFormat = "0.0" },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Kilograms", NumberFormat = "0.000\"kg\"" }
            }
        }));

    /// <summary>
    /// 验证来源工作表、缓存记录、数据字段格式编号及自定义格式定义。
    /// </summary>
    /// <param name="bytes">已导出的完整工作簿。</param>
    private static void AssertPackageFormatsAndCache(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var parts = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var partStream = entry.Open();
                return XDocument.Load(partStream);
            }, StringComparer.Ordinal);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var sharedStrings = parts.TryGetValue("xl/sharedStrings.xml", out var sharedStringsPart)
            ? sharedStringsPart.Root.Element(ns + "si").Parent.Elements(ns + "si")
                .Select(item => string.Concat(item.Descendants(ns + "t").Select(text => text.Value))).ToArray()
            : Array.Empty<string>();

        var workbook = parts["xl/workbook.xml"];
        var workbookRels = parts["xl/_rels/workbook.xml.rels"];
        var sheet = workbook.Root.Element(ns + "sheets").Elements(ns + "sheet")
            .Single(item => (string)item.Attribute("name") == "Data");
        var dataPart = ResolveRelationship("xl/workbook.xml", workbookRels,
            (string)sheet.Attribute(officeRel + "id"), parts);
        var dataSheet = parts[dataPart];
        Assert.Equal(new[]
        {
            "Category", "Amount", "OrderId",
            "North", "1.25", "PO-1",
            "South", "2.5", "PO-2",
            "North", "3.75", "PO-3"
        }, dataSheet.Root.Element(ns + "sheetData").Elements(ns + "row")
            .SelectMany(row => row.Elements(ns + "c").Select(cell => ReadCellValue(cell, ns, sharedStrings))));

        var pivotSheet = workbook.Root.Element(ns + "sheets").Elements(ns + "sheet")
            .Single(item => (string)item.Attribute("name") == "Summary");
        var pivotSheetPart = ResolveRelationship("xl/workbook.xml", workbookRels,
            (string)pivotSheet.Attribute(officeRel + "id"), parts);
        var pivotSheetRels = parts[RelationshipPartPath(pivotSheetPart)];
        var pivotTablePart = ResolveRelationship(pivotSheetPart, pivotSheetRels,
            pivotSheetRels.Root.Elements().Single(rel => ((string)rel.Attribute("Type")).EndsWith("pivotTable", StringComparison.Ordinal))
                .Attribute("Id").Value, parts);
        var pivotTable = parts[pivotTablePart].Root;
        Assert.Equal("FormatSummary", (string)pivotTable.Attribute("name"));

        var dataFields = pivotTable.Element(ns + "dataFields").Elements(ns + "dataField").ToArray();
        Assert.Equal(ExpectedCaptions, dataFields.Select(field => (string)field.Attribute("name")));
        Assert.All(dataFields, field =>
        {
            Assert.Equal("1", (string)field.Attribute("fld"));
            var subtotal = (string)field.Attribute("subtotal");
            Assert.True(subtotal == null || subtotal == "sum");
        });

        var styles = parts["xl/styles.xml"].Root;
        var customFormats = styles.Element(ns + "numFmts")?.Elements(ns + "numFmt")
            .ToDictionary(item => (string)item.Attribute("numFmtId"), item => (string)item.Attribute("formatCode"))
            ?? new Dictionary<string, string>();
        var formatIds = dataFields.Select(field => (string)field.Attribute("numFmtId")).ToArray();
        Assert.All(formatIds, Assert.NotNull);
        Assert.Equal("2", formatIds[0]);
        Assert.Equal("0.00", ResolveNumberFormat(formatIds[0], customFormats));
        Assert.Equal("0.0", ResolveNumberFormat(formatIds[1], customFormats));
        Assert.Equal("0.000\"kg\"", ResolveNumberFormat(formatIds[2], customFormats));
        Assert.Equal(3, formatIds.Distinct(StringComparer.Ordinal).Count());

        var cacheId = (string)pivotTable.Attribute("cacheId");
        var cacheEntry = workbook.Root.Element(ns + "pivotCaches").Elements(ns + "pivotCache")
            .Single(item => (string)item.Attribute("cacheId") == cacheId);
        var cachePart = ResolveRelationship("xl/workbook.xml", workbookRels,
            (string)cacheEntry.Attribute(officeRel + "id"), parts);
        var cache = parts[cachePart].Root;
        Assert.Equal("Data", (string)cache.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("sheet"));
        Assert.Equal("A1:C4", (string)cache.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref"));
        var cacheFields = cache.Element(ns + "cacheFields").Elements(ns + "cacheField").ToArray();
        Assert.Equal(new[] { "Category", "Amount", "OrderId" }, cacheFields.Select(field => (string)field.Attribute("name")));

        var cacheRels = parts[RelationshipPartPath(cachePart)];
        var recordsPart = ResolveRelationship(cachePart, cacheRels,
            cacheRels.Root.Elements().Single(rel => ((string)rel.Attribute("Type")).EndsWith("pivotCacheRecords", StringComparison.Ordinal))
                .Attribute("Id").Value, parts);
        var records = parts[recordsPart].Root.Elements(ns + "r").ToArray();
        Assert.Equal(3, records.Length);
        Assert.All(records, row => Assert.Equal(3, row.Elements().Count()));
        Assert.Equal(new[]
        {
            "North", "1.25", "PO-1",
            "South", "2.5", "PO-2",
            "North", "3.75", "PO-3"
        }, records.SelectMany(row => row.Elements().Select((cell, index) => ReadCacheValue(cell, cacheFields[index], ns))));
    }

    /// <summary>
    /// 解析工作簿包中的关系目标并确认目标部件存在。
    /// </summary>
    /// <param name="sourcePart">关系所属源部件路径。</param>
    /// <param name="relationships">源部件的关系文档。</param>
    /// <param name="relationshipId">关系标识。</param>
    /// <param name="parts">工作簿包中的 XML 部件。</param>
    /// <returns>规范化后的目标部件路径。</returns>
    private static string ResolveRelationship(string sourcePart, XDocument relationships, string relationshipId,
        IReadOnlyDictionary<string, XDocument> parts)
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
        Assert.Contains(normalized, parts.Keys);
        return normalized;
    }

    /// <summary>
    /// 获取指定源部件的关系文件路径。
    /// </summary>
    /// <param name="sourcePart">源部件路径。</param>
    /// <returns>对应关系文件路径。</returns>
    private static string RelationshipPartPath(string sourcePart)
    {
        var slash = sourcePart.LastIndexOf('/');
        return sourcePart.Substring(0, slash + 1) + "_rels/" + sourcePart.Substring(slash + 1) + ".rels";
    }

    /// <summary>
    /// 读取工作表 XML 单元格中的文本或数值。
    /// </summary>
    /// <param name="cell">工作表单元格节点。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <param name="sharedStrings">工作簿共享字符串表。</param>
    /// <returns>单元格的实际文本值。</returns>
    private static string ReadCellValue(XElement cell, XNamespace ns, IReadOnlyList<string> sharedStrings)
    {
        var type = (string)cell.Attribute("t");
        var value = (string)cell.Element(ns + "v");
        if (type == "inlineStr")
            return string.Concat(cell.Descendants(ns + "t").Select(item => item.Value));
        if (type == "s")
            return sharedStrings[int.Parse(value)];
        return value;
    }

    /// <summary>
    /// 读取透视缓存记录，并解析共享字符串索引。
    /// </summary>
    /// <param name="cell">缓存记录中的字段值节点。</param>
    /// <param name="field">对应缓存字段定义。</param>
    /// <param name="ns">SpreadsheetML 命名空间。</param>
    /// <returns>缓存中的逻辑值。</returns>
    private static string ReadCacheValue(XElement cell, XElement field, XNamespace ns)
    {
        var value = (string)cell.Attribute("v");
        if (cell.Name == ns + "x")
            return (string)field.Element(ns + "sharedItems").Elements().ElementAt(int.Parse(value)).Attribute("v");
        return value;
    }

    /// <summary>
    /// 将数字格式编号解析为 Excel 格式字符串。
    /// </summary>
    /// <param name="formatId">透视数据字段的格式编号。</param>
    /// <param name="customFormats">自定义格式编号映射。</param>
    /// <returns>编号对应的完整格式字符串。</returns>
    private static string ResolveNumberFormat(string formatId, IReadOnlyDictionary<string, string> customFormats)
    {
        if (customFormats.TryGetValue(formatId, out var customFormat))
            return customFormat;
        return formatId switch
        {
            "2" => "0.00",
            _ => null
        };
    }

    /// <summary>
    /// 读取 ClosedXML 透视值字段的内建或自定义数字格式。
    /// </summary>
    /// <param name="value">待读取格式的透视值字段。</param>
    /// <returns>格式编号对应的内建格式或自定义格式文本。</returns>
    private static string GetNumberFormat(IXLPivotValue value)
    {
        if (value.NumberFormat.NumberFormatId == 2)
            return "0.00";
        return value.NumberFormat.Format;
    }

}
