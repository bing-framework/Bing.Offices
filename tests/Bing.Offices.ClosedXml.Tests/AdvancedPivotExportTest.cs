using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证高级透视的真实工作簿结构与来源预检。
/// </summary>
public sealed class AdvancedPivotExportTest
{
    /// <summary>
    /// 验证多维透视保存全部字段配置与缓存记录。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出。</param>
    /// <param name="multiple">是否选择多个状态。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task MultiAxisAndValues_ShouldPreserveCompleteWorkbook(bool asynchronous, bool multiple)
    {
        var request = CreateRequest(multiple: multiple);
        Assert.True(new ClosedXmlExcelExporter().Supports(Bing.Offices.Providers.ExcelProviderFeatures.AdvancedPivotTables));
        using var output = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();
        if (asynchronous)
            await exporter.ExportAsync(request, output);
        else
            exporter.Export(request, output);
        Assert.True(output.CanWrite);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "Summary" }, workbook.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[]
        {
            "Warehouse", "Category", "Month", "Status", "Amount", "Item",
            "North", "Food", "Jan", "Open", "10", "A",
            "North", "Food", "Feb", "Closed", "30", "B",
            "North", "Tools", "Jan", "Open", "20", "C",
            "South", "Food", "Jan", "Pending", "40", "D"
        }, workbook.Worksheet("Data").Range("A1:F5").Cells().Select(cell => cell.GetString()));
        var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
        Assert.Equal("WarehouseReport", pivot.Name);
        Assert.Equal(new[] { "Warehouse", "Category" }, pivot.RowLabels.Select(field => field.SourceName));
        Assert.Equal(XLPivotSortType.Ascending, pivot.RowLabels.Get("Warehouse").SortType);
        Assert.Equal(XLPivotSortType.Descending, pivot.RowLabels.Get("Category").SortType);
        Assert.Equal(new[] { XLSubtotalFunction.Automatic }, pivot.RowLabels.Get("Warehouse").Subtotals);
        Assert.Empty(pivot.RowLabels.Get("Category").Subtotals);
        Assert.Contains(pivot.ColumnLabels, field => field.SourceName == "Month");
        Assert.Equal(new[] { "Total", "Mean", "Low", "High", "Count" },
            pivot.Values.Select(value => value.CustomName));
        Assert.Equal(new[] { XLPivotSummary.Sum, XLPivotSummary.Average, XLPivotSummary.Minimum,
            XLPivotSummary.Maximum, XLPivotSummary.Count }, pivot.Values.Select(value => value.SummaryFormula));
        var filter = Assert.Single(pivot.ReportFilters);
        Assert.Equal("Status", filter.SourceName);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()));
        var cache = ReadPart(archive, "pivotCacheDefinition");
        var records = ReadPart(archive, "pivotCacheRecords");
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var fields = cache.Root.Element(ns + "cacheFields").Elements().ToArray();
        Assert.Equal(new[] { "Warehouse", "Category", "Month", "Status", "Amount", "Item" },
            fields.Select(field => (string)field.Attribute("name")));
        var rows = records.Root.Elements(ns + "r").ToArray();
        Assert.Equal(4, rows.Length);
        Assert.All(rows, row => Assert.Equal(6, row.Elements().Count()));
        foreach (var row in rows)
            foreach (var fieldIndex in new[] { 0, 1, 2, 3, 5 })
            {
                var cell = row.Elements().ElementAt(fieldIndex);
                Assert.Equal("x", cell.Name.LocalName);
                var sharedItems = fields[fieldIndex].Element(ns + "sharedItems").Elements().ToArray();
                var sharedIndex = int.Parse((string)cell.Attribute("v"));
                Assert.InRange(sharedIndex, 0, sharedItems.Length - 1);
            }
        Assert.Equal(new[]
        {
            "North", "Food", "Jan", "Open", "10", "A",
            "North", "Food", "Feb", "Closed", "30", "B",
            "North", "Tools", "Jan", "Open", "20", "C",
            "South", "Food", "Jan", "Pending", "40", "D"
        }, rows.SelectMany(row => row.Elements().Select((cell, index) => cell.Name.LocalName == "x"
            ? (string)fields[index].Element(ns + "sharedItems").Elements()
                .ElementAt(int.Parse((string)cell.Attribute("v"))).Attribute("v")
            : (string)cell.Attribute("v"))));
        Assert.Single(cache.Root.Element(ns + "cacheSource").Elements());
        Assert.Equal("A1:F5", (string)cache.Root.Element(ns + "cacheSource")
            .Element(ns + "worksheetSource").Attribute("ref"));
        var pivotPart = ReadPart(archive, "pivotTable");
        var page = Assert.Single(pivotPart.Root.Element(ns + "pageFields").Elements());
        Assert.Equal("3", (string)page.Attribute("fld"));
        var pageItems = pivotPart.Root.Element(ns + "pivotFields").Elements().ElementAt(3)
            .Element(ns + "items").Elements().ToArray();
        var cacheItems = fields[3].Element(ns + "sharedItems").Elements().ToArray();
        // 单选用 pageField.item 表示；重开后的 SelectedValues 并不读取该索引。
        var selected = multiple
            ? pageItems.Where(item => item.Attribute("x") != null && (string)item.Attribute("h") != "1")
                .Select(item => int.Parse((string)item.Attribute("x"))).ToArray()
            : new[] { int.Parse((string)pageItems[int.Parse((string)page.Attribute("item"))].Attribute("x")) };
        Assert.Equal(multiple ? new[] { "Open", "Pending" } : new[] { "Open" },
            selected.Select(index => (string)cacheItems[index].Attribute("v")));
        if (multiple)
            Assert.Equal(new[] { "Open", "Pending" }, filter.SelectedValues.Select(value => value.ToString()));
        var dataFields = pivotPart.Root.Element(ns + "dataFields").Elements().ToArray();
        using var stylesStream = archive.GetEntry("xl/styles.xml").Open();
        var styles = XDocument.Load(stylesStream);
        var numberFormats = styles.Root.Element(ns + "numFmts")?.Elements()
            .ToDictionary(item => (string)item.Attribute("numFmtId"), item => (string)item.Attribute("formatCode"));
        var formatId = (string)dataFields[0].Attribute("numFmtId");
        Assert.True(formatId == "2" || formatId != null && numberFormats?.GetValueOrDefault(formatId) == "0.00",
            "Total 度量必须保留 0.00 数字格式。");
    }

    /// <summary>
    /// 验证非法筛选和展开预算在写入前失败。
    /// </summary>
    /// <param name="invalidFilter">是否指定来源中不存在的筛选值。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidFilterOrBudget_ShouldPreserveDestination(bool invalidFilter)
    {
        var request = CreateRequest(invalidFilter, invalidFilter ? null : 1);
        using var output = new MemoryStream();
        output.Write(new byte[] { 7, 11, 19, 23 });
        output.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(1, exception.RowIndex);
        Assert.Equal(invalidFilter ? 4 : 1, exception.ColumnIndex);
        Assert.Equal(new byte[] { 7, 11, 19, 23 }, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 创建仓库多维透视请求。
    /// </summary>
    /// <param name="invalidFilter">是否使用不存在的状态。</param>
    /// <param name="budget">可选完整展开面积上限。</param>
    /// <param name="multiple">是否选择多个状态。</param>
    /// <returns>包含全部来源行的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(bool invalidFilter = false, long? budget = null,
        bool multiple = false) =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new WarehouseRow { Warehouse = "North", Category = "Food", Month = "Jan", Status = "Open", Amount = 10, Item = "A" },
            new WarehouseRow { Warehouse = "North", Category = "Food", Month = "Feb", Status = "Closed", Amount = 30, Item = "B" },
            new WarehouseRow { Warehouse = "North", Category = "Tools", Month = "Jan", Status = "Open", Amount = 20, Item = "C" },
            new WarehouseRow { Warehouse = "South", Category = "Food", Month = "Jan", Status = "Pending", Amount = 40, Item = "D" }
        }).PivotTable(new ExcelPivotTableDefinition
        {
            Name = "WarehouseReport", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 4, EndColumn = 5 },
            RowFields = new[]
            {
                new ExcelPivotAxisFieldDefinition { Field = "Warehouse", SortOrder = ExcelPivotSortOrder.Ascending },
                new ExcelPivotAxisFieldDefinition { Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                    SubtotalMode = ExcelPivotSubtotalMode.None }
            },
            ColumnFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Month", SubtotalMode = ExcelPivotSubtotalMode.None } },
            ValueFields = new[]
            {
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Mean", Aggregation = ExcelPivotAggregation.Average },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Low", Aggregation = ExcelPivotAggregation.Minimum },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "High", Aggregation = ExcelPivotAggregation.Maximum },
                new ExcelPivotValueDefinition { Field = "Item", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
            },
            Filters = new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = multiple
                ? new object[] { "open", "Pending" } : new object[] { invalidFilter ? "Missing" : "open" } } },
            MaxExpandedCells = budget
        }));

    /// <summary>
    /// 读取唯一透视 XML 部件。
    /// </summary>
    /// <param name="archive">真实工作簿包。</param>
    /// <param name="prefix">部件文件名前缀。</param>
    /// <returns>完整部件文档。</returns>
    private static XDocument ReadPart(ZipArchive archive, string prefix)
    {
        var entry = Assert.Single(archive.Entries.Where(item =>
            item.FullName.Split('/').Last().StartsWith(prefix, StringComparison.Ordinal)
            && item.FullName.EndsWith(".xml", StringComparison.Ordinal)));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 仓库采购汇总来源行。
    /// </summary>
    private sealed class WarehouseRow
    {
        /// <summary>
        /// 初始化一个 <see cref="WarehouseRow"/> 类型的实例。
        /// </summary>
        public WarehouseRow() { }

        /// <summary>
        /// 获取或设置仓库。
        /// </summary>
        public string Warehouse { get; set; }

        /// <summary>
        /// 获取或设置商品分类。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置月份。
        /// </summary>
        public string Month { get; set; }

        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置商品。
        /// </summary>
        public string Item { get; set; }
    }
}
