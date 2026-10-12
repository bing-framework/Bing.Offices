using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证原生透视的列轴布局及多个缓存隔离。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotCompositionTest
{
    /// <summary>
    /// 原始业务字段按导出映射顺序排列。
    /// </summary>
    private static readonly string[] Headers = { "Region", "Amount", "State", "Channel", "Code" };

    /// <summary>
    /// 验证仅列轴的单度量或多度量保留两个筛选及完整来源。
    /// </summary>
    /// <param name="multipleValues">是否输出多个度量。</param>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Export_ColumnOnlyAndMultipleFilters_ShouldPreserveCompleteLayout(
        bool multipleValues, bool asynchronous)
    {
        var rows = Rows(0);
        var definition = Definition("ColumnOnly", "Summary", "Data", "Open", "0.00");
        definition = new ExcelPivotTableDefinition
        {
            Name = definition.Name, SourceSheet = definition.SourceSheet,
            TargetSheet = definition.TargetSheet, SourceRange = definition.SourceRange,
            TargetRow = 2, TargetColumn = 2,
            ColumnFields = new[] { new ExcelPivotAxisFieldDefinition
                { Field = "Channel", SubtotalMode = ExcelPivotSubtotalMode.None } },
            ValueFields = multipleValues ? new[]
            {
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" },
                new ExcelPivotValueDefinition { Field = "Code", Caption = "CodeCount", Aggregation = ExcelPivotAggregation.Count }
            } : new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" } },
            Filters = new[]
            {
                new ExcelPivotFilterDefinition { Field = "State", Values = new object[] { "Open" } },
                new ExcelPivotFilterDefinition { Field = "Region", Values = new object[] { "North" } }
            }
        };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows).PivotTable(definition));
        var bytes = await Export(request, asynchronous);
        using var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        AssertSource(workbook, "Data", rows);
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var pivot = ((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables().Single();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var table = ReadXml(archive, pivot.GetPackagePart().PartName.Name).Root;
        Assert.Null(table.Element(ns + "rowFields"));
        Assert.Equal(multipleValues ? new[] { 3, -2 } : new[] { 3 },
            table.Element(ns + "colFields").Elements().Select(field => (int)field.Attribute("x")));
        Assert.Equal(new[] { 2, 0 }, table.Element(ns + "pageFields").Elements()
            .Select(field => (int)field.Attribute("fld")));
        Assert.All(table.Element(ns + "pageFields").Elements(), field => Assert.Equal(0u, (uint)field.Attribute("item")));
        Assert.Equal(multipleValues ? new[] { "sum", "count" } : new[] { "sum" },
            table.Element(ns + "dataFields").Elements().Select(field => (string)field.Attribute("subtotal") ?? "sum"));
        var location = table.Element(ns + "location");
        Assert.Equal(multipleValues ? "C6:I9" : "C6:F8", (string)location.Attribute("ref"));
        Assert.Equal(2u, (uint)location.Attribute("rowPageCount"));
        Assert.Equal(1u, (uint)location.Attribute("colPageCount"));
        AssertCache(archive, pivot, rows);
    }

    /// <summary>
    /// 验证同源和异源透视的筛选、格式及记录在重复导出时独立。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_MultiplePivotsAndRepeat_ShouldKeepAllCachesIsolated(bool asynchronous)
    {
        var rows = Rows(0);
        var otherRows = Rows(100);
        var definitions = new[]
        {
            Definition("FirstPivot", "First", "Data", "Open", "0.00"),
            Definition("SecondPivot", "Second", "Data", "Closed", "0.000"),
            Definition("ThirdPivot", "Third", "Other", "Open", "0.0")
        };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows).AddSheet("Other", otherRows)
            .PivotTable(definitions[0]).PivotTable(definitions[1]).PivotTable(definitions[2]));
        for (var iteration = 0; iteration < 2; iteration++)
        {
            var bytes = await Export(request, asynchronous);
            using var workbook = new XSSFWorkbook(new MemoryStream(bytes));
            using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            Assert.Equal(new[] { "Data", "Other", "First", "Second", "Third" },
                Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
            AssertSource(workbook, "Data", rows);
            AssertSource(workbook, "Other", otherRows);
            var cacheParts = new HashSet<string>();
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            foreach (var definition in definitions)
            {
                var pivot = ((XSSFSheet)workbook.GetSheet(definition.TargetSheet)).GetPivotTables().Single();
                Assert.True(cacheParts.Add(CachePart(archive, pivot)));
                var table = ReadXml(archive, pivot.GetPackagePart().PartName.Name).Root;
                Assert.Equal(definition.Name, (string)table.Attribute("name"));
                var selection = table.Element(ns + "pageFields").Elements().Single();
                Assert.Equal(definition.TargetSheet == "Second" ? 1u : 0u, (uint)selection.Attribute("item"));
                var formatId = (short)(uint)table.Element(ns + "dataFields").Elements().Single().Attribute("numFmtId");
                Assert.Equal(definition.ValueFields[0].NumberFormat, workbook.CreateDataFormat().GetFormat(formatId));
                AssertCache(archive, pivot, definition.SourceSheet == "Data" ? rows : otherRows);
            }
            Assert.Equal(3, cacheParts.Count);
        }
    }

    /// <summary>
    /// 验证后续透视预检失败不写出目标且不污染复用请求。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_LaterInvalidDefinition_ShouldPreserveDestinationAndRetry(bool asynchronous)
    {
        var first = Definition("FirstPivot", "First", "Data", "Open", "0.00");
        var invalid = new ExcelPivotTableDefinition
        {
            Name = "InvalidPivot", SourceSheet = "Data", TargetSheet = "Second",
            SourceRange = first.SourceRange, RowField = "Region", ValueField = "Missing"
        };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", Rows(0)).PivotTable(first).PivotTable(invalid));
        for (var iteration = 0; iteration < 2; iteration++)
        {
            using var output = new MemoryStream();
            var original = new byte[] { 11, 22, 33, 44 };
            output.Write(original);
            var exception = await Assert.ThrowsAsync<BingOfficesConfigurationException>(async () =>
            {
                if (asynchronous) await new NpoiExcelExporter().ExportAsync(request, output);
                else new NpoiExcelExporter().Export(request, output);
            });
            Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
            Assert.Contains("Missing", exception.Message);
            Assert.Equal(original, output.ToArray());
            Assert.True(output.CanWrite);
            Assert.Equal(new[] { "FirstPivot", "InvalidPivot" }, request.PivotTables.Select(pivot => pivot.Name));
        }
    }

    /// <summary>
    /// 导出完整工作簿并核对调用方流所有权。
    /// </summary>
    /// <param name="request">包含透视配置的不可变请求。</param>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    /// <returns>完整的工作簿内容。</returns>
    private static async Task<byte[]> Export(ExcelWorkbookExportRequest request, bool asynchronous)
    {
        using var output = new MemoryStream();
        if (asynchronous) await new NpoiExcelExporter().ExportAsync(request, output);
        else new NpoiExcelExporter().Export(request, output);
        Assert.True(output.CanWrite);
        return output.ToArray();
    }

    /// <summary>
    /// 核对全部业务字段及记录。
    /// </summary>
    /// <param name="workbook">实际重开的工作簿。</param>
    /// <param name="name">来源工作表名称。</param>
    /// <param name="rows">期望的完整记录。</param>
    private static void AssertSource(XSSFWorkbook workbook, string name, IReadOnlyList<NpoiPivotExportRow> rows)
    {
        var sheet = workbook.GetSheet(name);
        Assert.Equal(rows.Count, sheet.LastRowNum);
        Assert.Equal(Headers, Enumerable.Range(0, Headers.Length).Select(column => sheet.GetRow(0).GetCell(column).StringCellValue));
        for (var row = 0; row < rows.Count; row++)
        {
            var cells = sheet.GetRow(row + 1);
            Assert.Equal(rows[row].Region, cells.GetCell(0).StringCellValue);
            Assert.Equal(rows[row].Amount, cells.GetCell(1).NumericCellValue);
            Assert.Equal(rows[row].State, cells.GetCell(2).StringCellValue);
            Assert.Equal(rows[row].Channel, cells.GetCell(3).StringCellValue);
            Assert.Equal(rows[row].Code, cells.GetCell(4).StringCellValue);
        }
    }

    /// <summary>
    /// 按真实缓存关系核对全部共享项引用与来源值。
    /// </summary>
    /// <param name="archive">实际导出的工作簿包。</param>
    /// <param name="pivot">公开读取的原生透视。</param>
    /// <param name="rows">期望的完整来源记录。</param>
    private static void AssertCache(ZipArchive archive, XSSFPivotTable pivot, IReadOnlyList<NpoiPivotExportRow> rows)
    {
        var part = CachePart(archive, pivot);
        var definition = ReadXml(archive, part).Root;
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRelations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var target = ResolveRelatedPart(archive, part, (string)definition.Attribute(officeRelations + "id"), "pivotCacheRecords");
        var records = ReadXml(archive, target).Root;
        Assert.Equal(rows.Count, (int)definition.Attribute("recordCount"));
        Assert.Equal(rows.Count, (int)records.Attribute("count"));
        Assert.Equal("A1:E4", (string)definition.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref"));
        var fields = definition.Element(ns + "cacheFields").Elements().ToArray();
        Assert.Equal(Headers, fields.Select(field => (string)field.Attribute("name")));
        var recordRows = records.Elements().ToArray();
        Assert.Equal(rows.Count, recordRows.Length);
        for (var row = 0; row < rows.Count; row++)
        {
            var indexes = recordRows[row].Elements().Select(item => (int)item.Attribute("v")).ToArray();
            Assert.Equal(5, indexes.Length);
            var shared = fields.Select((field, column) => field.Element(ns + "sharedItems").Elements().ElementAt(indexes[column])).ToArray();
            Assert.Equal(rows[row].Region, (string)shared[0].Attribute("v"));
            Assert.Equal(rows[row].Amount, (double)shared[1].Attribute("v"));
            Assert.Equal(rows[row].State, (string)shared[2].Attribute("v"));
            Assert.Equal(rows[row].Channel, (string)shared[3].Attribute("v"));
            Assert.Equal(rows[row].Code, (string)shared[4].Attribute("v"));
        }
    }

    /// <summary>
    /// 读取包内指定的 XML 部件。
    /// </summary>
    /// <param name="archive">实际工作簿包。</param>
    /// <param name="name">部件的完整地址。</param>
    /// <returns>解析后的完整 XML。</returns>
    private static XDocument ReadXml(ZipArchive archive, string name)
    {
        using var stream = archive.GetEntry(name.TrimStart('/')).Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 按工作簿缓存标识及关系解析真实缓存定义地址。
    /// </summary>
    /// <param name="archive">实际工作簿包。</param>
    /// <param name="pivot">公开重开的透视定义。</param>
    /// <returns>真实缓存定义部件地址。</returns>
    /// <remarks>
    /// 原生重开对象不一定建立缓存 getter；验证使用实际 OPC 关系而不是猜测部件编号。
    /// </remarks>
    private static string CachePart(ZipArchive archive, XSSFPivotTable pivot)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace relations = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var table = ReadXml(archive, pivot.GetPackagePart().PartName.Name).Root;
        var cache = ReadXml(archive, "xl/workbook.xml").Root.Element(ns + "pivotCaches").Elements()
            .Single(item => (uint)item.Attribute("cacheId") == (uint)table.Attribute("cacheId"));
        return ResolveRelatedPart(archive, "xl/workbook.xml", (string)cache.Attribute(relations + "id"), "pivotCacheDefinition");
    }

    /// <summary>
    /// 解析并核对内部 OPC 关系目标。
    /// </summary>
    /// <param name="archive">实际工作簿包。</param>
    /// <param name="part">拥有关系的部件地址。</param>
    /// <param name="id">待解析的关系标识。</param>
    /// <param name="type">关系类型的末尾名称。</param>
    /// <returns>规范化后的包内目标地址。</returns>
    private static string ResolveRelatedPart(ZipArchive archive, string part, string id, string type)
    {
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        var relationPart = part[..part.LastIndexOf('/')] + "/_rels/" + part[(part.LastIndexOf('/') + 1)..] + ".rels";
        var relation = ReadXml(archive, relationPart).Root.Elements(ns + "Relationship")
            .Single(item => (string)item.Attribute("Id") == id);
        Assert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/" + type,
            (string)relation.Attribute("Type"));
        Assert.NotEqual("External", (string)relation.Attribute("TargetMode"));
        var target = new Uri(new Uri("https://package.local/" + part), (string)relation.Attribute("Target")).AbsolutePath.TrimStart('/');
        Assert.NotNull(archive.GetEntry(target));
        return target;
    }

    /// <summary>
    /// 创建具有独立筛选和数字格式的透视定义。
    /// </summary>
    /// <param name="name">透视名称。</param>
    /// <param name="target">目标工作表名称。</param>
    /// <param name="source">来源工作表名称。</param>
    /// <param name="state">选中的业务状态。</param>
    /// <param name="format">度量数字格式。</param>
    /// <returns>完整的透视定义。</returns>
    private static ExcelPivotTableDefinition Definition(string name, string target, string source, string state, string format) => new()
    {
        Name = name, TargetSheet = target, SourceSheet = source,
        SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 4 },
        RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
        ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = name + "Total", NumberFormat = format } },
        Filters = new[] { new ExcelPivotFilterDefinition { Field = "State", Values = new object[] { state } } }
    };

    /// <summary>
    /// 创建可区分来源工作表的完整记录。
    /// </summary>
    /// <param name="offset">来源金额偏移量。</param>
    /// <returns>包含筛选外记录的完整业务数据。</returns>
    private static NpoiPivotExportRow[] Rows(double offset) => new[]
    {
        new NpoiPivotExportRow { Region = "North", Amount = offset + 10, State = "Open", Channel = "Retail", Code = "A" },
        new NpoiPivotExportRow { Region = "North", Amount = offset + 20, State = "Closed", Channel = "Wholesale", Code = "B" },
        new NpoiPivotExportRow { Region = "South", Amount = offset + 30, State = "Open", Channel = "Retail", Code = "C" }
    };
}
