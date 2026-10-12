using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证原生透视字段、完整缓存及创建隔离。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotTableWriterTest
{
    /// <summary>
    /// 验证多轴五聚合、数字格式和多选筛选保留全部来源。
    /// </summary>
    [Fact]
    public void Create_ShouldPersistAllRolesAndCompleteUnfilteredCache()
    {
        using var workbook = CreateSource();
        var definition = Advanced(new object[] { "open", "Closed" });
        var source = NpoiPivotSourceValidator.Validate(workbook, definition, default);

        var records = NpoiPivotTableWriter.Create(workbook, source, default);

        var pivot = ((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables().Single();
        var table = pivot.GetCTPivotTableDefinition();
        Assert.Equal("PurchasePivot", table.name);
        Assert.Equal(new[] { 0, 1 }, table.rowFields.field.Select(field => field.x));
        Assert.Equal(new[] { 2, -2 }, table.colFields.field.Select(field => field.x));
        Assert.Equal(new uint[] { 3, 4, 3, 3, 3 }, table.dataFields.dataField.Select(field => field.fld));
        Assert.Equal(new[] { ST_DataConsolidateFunction.sum, ST_DataConsolidateFunction.count,
            ST_DataConsolidateFunction.average, ST_DataConsolidateFunction.min, ST_DataConsolidateFunction.max },
            table.dataFields.dataField.Select(field => field.subtotal));
        Assert.All(table.dataFields.dataField, field =>
        {
            Assert.Equal(-1, field.baseField);
            Assert.Equal(1048832u, field.baseItem);
        });
        Assert.Equal(new[] { "Total", "CodeCount", "Mean", "Minimum", "Maximum" },
            table.dataFields.dataField.Select(field => field.name));
        Assert.True(table.dataFields.dataField[0].numFmtIdSpecified);
        Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat((short)table.dataFields.dataField[0].numFmtId));
        Assert.False(table.compact);
        Assert.False(table.outline);
        Assert.True(table.rowGrandTotals);
        Assert.True(table.colGrandTotals);
        Assert.Equal(ST_FieldSortType.ascending, table.pivotFields.pivotField[0].sortType);
        Assert.True(table.pivotFields.pivotField[0].defaultSubtotal);
        Assert.False(table.pivotFields.pivotField[1].defaultSubtotal);
        Assert.Equal(ST_FieldSortType.descending, table.pivotFields.pivotField[2].sortType);
        Assert.Equal(ST_Axis.axisCol, table.pivotFields.pivotField[2].axis);
        Assert.Equal(ST_Axis.axisPage, table.pivotFields.pivotField[5].axis);
        Assert.True(table.pivotFields.pivotField[5].multipleItemSelectionAllowed);
        Assert.Equal(new[] { false, false, true }, table.pivotFields.pivotField[5].items.item.Select(item => item.h));
        Assert.False(table.pivotFields.pivotField[4].countSubtotal);
        Assert.False(table.pageFields.pageField.Single().itemSpecified);
        Assert.Equal("C6:S13", table.location.@ref);
        Assert.Equal(2u, table.location.firstDataCol);
        Assert.Equal(2u, table.location.firstDataRow);
        Assert.Equal(3, records.Rows.Count);
        Assert.All(records.Rows, row => Assert.Equal(7, row.Count));
        using var output = new MemoryStream();
        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(), new[] { records }, output, default);
        Assert.True(output.CanWrite);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var pivotPart = archive.Entries.Single(entry =>
            entry.FullName.StartsWith("xl/pivotTables/pivotTable", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var pivotStream = pivotPart.Open();
        var pivotXml = XDocument.Load(pivotStream);
        var dataFieldsXml = pivotXml.Root.Element(ns + "dataFields").Elements(ns + "dataField").ToArray();
        Assert.All(dataFieldsXml, field =>
        {
            Assert.Equal("normal", (string)field.Attribute("showDataAs"));
            Assert.Equal("-1", (string)field.Attribute("baseField"));
            Assert.Equal("1048832", (string)field.Attribute("baseItem"));
        });
        var countDataFieldXml = dataFieldsXml.Single(field =>
            (string)field.Attribute("name") == "CodeCount");
        Assert.Equal("count", (string)countDataFieldXml.Attribute("subtotal"));
        var cachePart = archive.Entries.Single(entry =>
            entry.FullName.StartsWith("xl/pivotCache/pivotCacheDefinition", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var cacheStream = cachePart.Open();
        var cacheXml = XDocument.Load(cacheStream);
        var codeSharedItems = cacheXml.Root.Element(ns + "cacheFields")
            .Elements(ns + "cacheField").ElementAt(4).Element(ns + "sharedItems");
        Assert.Equal("1", (string)codeSharedItems.Attribute("containsString"));
        Assert.Equal("1", (string)codeSharedItems.Attribute("containsSemiMixedTypes"));
        Assert.Equal("1", (string)codeSharedItems.Attribute("containsNonDate"));
        var part = archive.GetEntry(records.PartName);
        using var stream = part.Open();
        var xml = XDocument.Load(stream);
        Assert.Equal(3, (int)xml.Root.Attribute("count"));
        Assert.Equal(3, xml.Root.Elements(ns + "r").Count());
        for (var row = 0; row < records.Rows.Count; row++)
            Assert.Equal(records.Rows[row], xml.Root.Elements(ns + "r").ElementAt(row)
                .Elements(ns + "x").Select(item => (uint)item.Attribute("v")));
        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "Summary" }, Enumerable.Range(0, reopened.NumberOfSheets)
            .Select(reopened.GetSheetName));
        Assert.Equal("Ignored", reopened.GetSheet("Data").GetRow(3).GetCell(5).StringCellValue);
        Assert.Equal("unused-3", reopened.GetSheet("Data").GetRow(3).GetCell(6).StringCellValue);
        Assert.Equal(30d, reopened.GetSheet("Data").GetRow(3).GetCell(3).NumericCellValue);
        Assert.Single(((XSSFSheet)reopened.GetSheet("Summary")).GetPivotTables());
    }

    /// <summary>
    /// 记录 NPOI 原生 Count 度量字段的形状，作为自写透视 XML 的对照。
    /// </summary>
    [Fact]
    public void NativeCountField_ShouldExposeExpectedPivotFieldShape()
    {
        using var workbook = CreateNativeCountWorkbook(out var pivot);
        var table = pivot.GetCTPivotTableDefinition();
        var codeField = table.pivotFields.pivotField[1];
        var dataField = Assert.Single(table.dataFields.dataField);

        Assert.True(codeField.dataField);
        Assert.Equal(ST_DataConsolidateFunction.count, dataField.subtotal);
        Assert.Equal((uint)1, dataField.fld);
        Assert.Equal("CodeCount", dataField.name);
        Assert.False(codeField.countSubtotal);
        Assert.Null(codeField.items);
    }

    /// <summary>
    /// 验证单选与全部选只改变筛选项目而不裁剪记录。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Create_FilterSelection_ShouldKeepAllRecords(bool selected)
    {
        using var workbook = CreateSource();
        var source = NpoiPivotSourceValidator.Validate(workbook,
            Advanced(selected ? new object[] { "closed" } : Array.Empty<object>()), default);
        var records = NpoiPivotTableWriter.Create(workbook, source, default);
        var pivot = ((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables().Single();
        var table = pivot.GetCTPivotTableDefinition();
        var page = table.pageFields.pageField.Single();
        Assert.Equal(selected, page.itemSpecified);
        if (selected)
            Assert.Equal(1u, page.item);
        Assert.False(table.pivotFields.pivotField[5].multipleItemSelectionAllowed);
        Assert.All(table.pivotFields.pivotField[5].items.item, item => Assert.False(item.h));
        Assert.Equal(3, records.Rows.Count);
    }

    /// <summary>
    /// 验证独立工作簿并发创建和预取消不会污染目标。
    /// </summary>
    [Fact]
    public async Task Create_ConcurrentWorkbooksAndPreCancellation_ShouldRemainIsolated()
    {
        var tasks = Enumerable.Range(0, 8).Select(index => Task.Run(() =>
        {
            using var workbook = CreateSource();
            var source = NpoiPivotSourceValidator.Validate(workbook, Advanced(Array.Empty<object>()), default);
            var records = NpoiPivotTableWriter.Create(workbook, source, default);
            Assert.Equal(3, records.Rows.Count);
            Assert.Equal(2, workbook.NumberOfSheets);
            return records.Rows.Select(row => row.ToArray()).ToArray();
        })).ToArray();
        var results = await Task.WhenAll(tasks);
        foreach (var result in results)
            for (var row = 0; row < result.Length; row++)
                Assert.Equal(results[0][row], result[row]);
        using var canceledWorkbook = CreateSource();
        var canceledSource = NpoiPivotSourceValidator.Validate(canceledWorkbook, Advanced(Array.Empty<object>()), default);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var error = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotTableWriter.Create(canceledWorkbook, canceledSource, cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Equal(1, canceledWorkbook.NumberOfSheets);
    }

    /// <summary>
    /// 创建保留原始字段顺序的采购来源。
    /// </summary>
    /// <returns>包含全部业务数据的独立工作簿。</returns>
    private static XSSFWorkbook CreateSource()
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var headers = new[] { "Warehouse", "Category", "Month", "Amount", "Code", "State", "Other" };
        var header = sheet.CreateRow(0);
        for (var column = 0; column < headers.Length; column++)
            header.CreateCell(column).SetCellValue(headers[column]);
        for (var row = 0; row < 3; row++)
        {
            var cells = sheet.CreateRow(row + 1);
            cells.CreateCell(0).SetCellValue(row == 2 ? "South" : "North");
            cells.CreateCell(1).SetCellValue(row == 1 ? "Gift" : "Product");
            cells.CreateCell(2).SetCellValue(row == 1 ? "February" : "January");
            cells.CreateCell(3).SetCellValue((row + 1) * 10d);
            cells.CreateCell(4).SetCellValue("CODE-" + row);
            cells.CreateCell(5).SetCellValue(new[] { "Open", "Closed", "Ignored" }[row]);
            cells.CreateCell(6).SetCellValue("unused-" + (row + 1));
        }
        return workbook;
    }

    /// <summary>
    /// 创建由 NPOI 原生 API 配置文本 Count 的最小透视表。
    /// </summary>
    /// <param name="pivot">返回原生透视表。</param>
    /// <returns>包含文本来源字段的工作簿。</returns>
    private static XSSFWorkbook CreateNativeCountWorkbook(out XSSFPivotTable pivot)
    {
        var workbook = new XSSFWorkbook();
        var data = workbook.CreateSheet("Data");
        var summary = (XSSFSheet)workbook.CreateSheet("Summary");
        data.CreateRow(0).CreateCell(0, CellType.String).SetCellValue("Group");
        data.GetRow(0).CreateCell(1, CellType.String).SetCellValue("Code");
        var rows = new[] { new[] { "North", "A" }, new[] { "North", "B" }, new[] { "South", "A" } };
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = data.CreateRow(rowIndex + 1);
            row.CreateCell(0, CellType.String).SetCellValue(rows[rowIndex][0]);
            row.CreateCell(1, CellType.String).SetCellValue(rows[rowIndex][1]);
        }

        var source = new AreaReference(new CellReference("A1"),
            new CellReference(rows.Length, 1), SpreadsheetVersion.EXCEL2007);
        pivot = summary.CreatePivotTable(source, new CellReference("A1"), data);
        pivot.AddRowLabel(0);
        pivot.AddColumnLabel(DataConsolidateFunction.COUNT, 1, "CodeCount");
        return workbook;
    }

    /// <summary>
    /// 创建多轴五聚合采购透视定义。
    /// </summary>
    /// <param name="selection">状态筛选选项。</param>
    /// <returns>保留全部来源区域的高级定义。</returns>
    private static ExcelPivotTableDefinition Advanced(IReadOnlyList<object> selection) => new()
    {
        Name = "PurchasePivot", SourceSheet = "Data", TargetSheet = "Summary", TargetRow = 3, TargetColumn = 2,
        SourceRange = new ExcelRangeDefinition { StartRow = 0, EndRow = 3, StartColumn = 0, EndColumn = 6 },
        RowFields = new[]
        {
            new ExcelPivotAxisFieldDefinition { Field = "Warehouse", SortOrder = ExcelPivotSortOrder.Ascending },
            new ExcelPivotAxisFieldDefinition { Field = "Category", SubtotalMode = ExcelPivotSubtotalMode.None }
        },
        ColumnFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Month", SortOrder = ExcelPivotSortOrder.Descending } },
        ValueFields = new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" },
            new ExcelPivotValueDefinition { Field = "Code", Caption = "CodeCount", Aggregation = ExcelPivotAggregation.Count },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Mean", Aggregation = ExcelPivotAggregation.Average },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Minimum", Aggregation = ExcelPivotAggregation.Minimum },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Maximum", Aggregation = ExcelPivotAggregation.Maximum }
        },
        Filters = new[] { new ExcelPivotFilterDefinition { Field = "State", Values = selection } }
    };
}
