using System;
using System.IO;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证原生透视公开导出与完整来源缓存。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotPublicExportTest
{
    /// <summary>
    /// 验证同步和异步入口生成真实基础或高级透视。
    /// </summary>
    /// <param name="advanced">是否使用高级多度量模式。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Export_ShouldPersistCompleteSourceAndNativePivot(bool advanced, bool asynchronous)
    {
        var request = Request(advanced);
        var exporter = new NpoiExcelExporter();
        Assert.True(exporter.Supports(ExcelProviderFeatures.PivotTables | ExcelProviderFeatures.AdvancedPivotTables));
        using var output = new MemoryStream();
        if (asynchronous)
            await exporter.ExportAsync(request, output);
        else
            exporter.Export(request, output);

        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        AssertWorkbook(output.ToArray(), advanced);
        using var repeated = new MemoryStream();
        exporter.Export(request, repeated);
        AssertWorkbook(repeated.ToArray(), advanced);
        Assert.Equal(3, request.Sheets.Single().Data.Cast<NpoiPivotExportRow>().Count());
    }

    /// <summary>
    /// 验证来源预检失败和预取消均保全已有目标文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportToFile_FailureAndCancellation_ShouldPreserveTarget(bool asynchronous)
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-pivot-public-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "purchase.xlsx");
        var original = new byte[] { 17, 42, 91, 23 };
        File.WriteAllBytes(path, original);
        try
        {
            var exporter = new NpoiExcelExporter();
            var request = Request(false, endRow: 4);
            var error = asynchronous
                ? await Assert.ThrowsAsync<BingOfficesConfigurationException>(() => exporter.ExportToFileAsync(request, path))
                : Assert.Throws<BingOfficesConfigurationException>(() => exporter.ExportToFile(request, path));
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
            Assert.Equal("Data", error.SheetName);
            Assert.Equal(5, error.RowIndex);
            Assert.Equal(1, error.ColumnIndex);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            if (asynchronous)
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    exporter.ExportToFileAsync(Request(false), path, cancellation.Token));
            else
                Assert.Throws<OperationCanceledException>(() => exporter.ExportToFile(Request(false), path, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    /// <summary>
    /// 验证字典辅助工作表避开透视目标名称。
    /// </summary>
    [Fact]
    public void Export_WithLookupReservedTarget_ShouldKeepDictionaryAndPivot()
    {
        var request = ExcelExport.Workbook(builder => builder
            .AddSheet("Data", Rows())
            .LookupList("StatusValues", new[] { "Open", "Closed" })
            .PivotTable(Definition(false, 3, "__BingLookup_001")));
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "__BingLookup_002", "__BingLookup_001" },
            Enumerable.Range(0, reopened.NumberOfSheets).Select(reopened.GetSheetName));
        Assert.Equal(NPOI.SS.UserModel.SheetVisibility.Hidden, reopened.GetSheetVisibility(1));
        Assert.Equal(new[] { "Open", "Closed" }, Enumerable.Range(0, 2).Select(index =>
            reopened.GetSheetAt(1).GetRow(index).GetCell(0).StringCellValue));
        Assert.Equal("__BingLookup_002!$A$1:$A$2", reopened.GetName("StatusValues").RefersToFormula);
        Assert.Single(((XSSFSheet)reopened.GetSheetAt(2)).GetPivotTables());
        AssertWorkbook(output.ToArray(), false, "__BingLookup_001",
            new[] { "Data", "__BingLookup_002", "__BingLookup_001" });
    }

    /// <summary>
    /// 验证级联辅助工作表避开透视目标并保留完整字典与引用。
    /// </summary>
    [Fact]
    public void Export_WithCascadeReservedTarget_ShouldKeepCompleteDictionaryAndValidation()
    {
        var request = ExcelExport.Workbook(builder => builder
            .CascadeList("Channels", new Dictionary<string, IReadOnlyList<string>>
            {
                ["North"] = new[] { "Retail", "Wholesale" },
                ["South"] = new[] { "Retail" }
            })
            .AddSheet("Data", Rows(), sheet => sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Type = ExcelDataValidationType.List, CascadeSourceName = "Channels", ParentColumnIndex = 0,
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 3, StartColumn = 3, EndColumn = 3 }
            }))
            .PivotTable(Definition(false, 3, "__BingCascade_001")));
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "__BingCascade_002", "__BingCascade_001" },
            Enumerable.Range(0, reopened.NumberOfSheets).Select(reopened.GetSheetName));
        Assert.Equal(NPOI.SS.UserModel.SheetVisibility.Hidden, reopened.GetSheetVisibility(1));
        var helper = reopened.GetSheetAt(1);
        Assert.Equal(new[] { "North", "South" }, Enumerable.Range(0, 2).Select(index => helper.GetRow(index).GetCell(0).StringCellValue));
        Assert.Equal(new[] { "__BingCascadeValues_001", "__BingCascadeValues_002" },
            Enumerable.Range(0, 2).Select(index => helper.GetRow(index).GetCell(1).StringCellValue));
        Assert.Equal(new[] { "Retail", "Wholesale", "Retail" },
            Enumerable.Range(0, 3).Select(index => helper.GetRow(index).GetCell(2).StringCellValue));
        Assert.Equal("__BingCascade_002!$A$1:$A$2", reopened.GetName("Channels").RefersToFormula);
        Assert.Equal("__BingCascade_002!$A$1:$B$2", reopened.GetName("__BingCascadeMap_001").RefersToFormula);
        Assert.Equal("__BingCascade_002!$C$1:$C$2", reopened.GetName("__BingCascadeValues_001").RefersToFormula);
        Assert.Equal("__BingCascade_002!$C$3:$C$3", reopened.GetName("__BingCascadeValues_002").RefersToFormula);
        Assert.Equal("INDIRECT(VLOOKUP($A2,__BingCascadeMap_001,2,FALSE))",
            Assert.Single(reopened.GetSheet("Data").GetDataValidations()).ValidationConstraint.Formula1);
        Assert.Single(((XSSFSheet)reopened.GetSheetAt(2)).GetPivotTables());
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        Assert.Equal(3, ReadXml(archive, "pivotCacheRecords").Root.Elements().Count());
        AssertWorkbook(output.ToArray(), false, "__BingCascade_001",
            new[] { "Data", "__BingCascade_002", "__BingCascade_001" });
    }

    /// <summary>
    /// 验证普通模板的未覆盖样式、尺寸、合并、名称及公式仍保留。
    /// </summary>
    [Fact]
    public void Export_WithOrdinaryTemplate_ShouldPreserveUnrelatedMetadataAndFullCache()
    {
        using var template = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            var note = sheet.CreateRow(8).CreateCell(7);
            note.SetCellValue("Template note");
            var style = workbook.CreateCellStyle();
            style.IsLocked = false;
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00");
            note.CellStyle = style;
            sheet.GetRow(8).HeightInPoints = 28;
            sheet.SetColumnWidth(7, 22 * 256);
            sheet.CreateRow(9).CreateCell(7).SetCellValue("Approved");
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(9, 9, 7, 8));
            sheet.CreateRow(10).CreateCell(7).SetCellFormula("SUM(B2:B4)");
            var name = workbook.CreateName();
            name.NameName = "TemplateNote";
            name.RefersToFormula = "Data!$H$9";
            workbook.Write(template, true);
        }
        template.Position = 0;
        var original = template.ToArray();
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", Rows()).PivotTable(Definition(true, 3)));
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        Assert.True(template.CanRead);
        Assert.Equal(original, template.ToArray());
        AssertWorkbook(output.ToArray(), true);
        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        var data = reopened.GetSheet("Data");
        Assert.Equal("Template note", data.GetRow(8).GetCell(7).StringCellValue);
        Assert.False(data.GetRow(8).GetCell(7).CellStyle.IsLocked);
        Assert.Equal("0.00", data.GetRow(8).GetCell(7).CellStyle.GetDataFormatString());
        Assert.Equal(28, data.GetRow(8).HeightInPoints);
        Assert.Equal(22 * 256, data.GetColumnWidth(7));
        Assert.Equal("H10:I10", Assert.Single(data.MergedRegions).FormatAsString());
        Assert.Equal("Approved", data.GetRow(9).GetCell(7).StringCellValue);
        Assert.Equal("SUM(B2:B4)", data.GetRow(10).GetCell(7).CellFormula);
        Assert.Equal("Data!$H$9", reopened.GetName("TemplateNote").RefersToFormula);
    }

    /// <summary>
    /// 验证已有透视模板明确拒绝且保留调用方的输入输出内容。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_ExistingPivotTemplate_ShouldRejectAndPreserveStreams(bool asynchronous)
    {
        using var template = new MemoryStream();
        var exporter = new NpoiExcelExporter();
        exporter.Export(Request(false), template);
        var original = template.ToArray();
        template.Position = 0;
        var request = ExcelExport.Workbook(builder => builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", Rows()).PivotTable(Definition(true, 3, "NewSummary")));
        using var output = new MemoryStream();
        var previousOutput = new byte[] { 17, 42, 91 };
        output.Write(previousOutput);
        var error = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() => exporter.ExportAsync(request, output))
            : Assert.Throws<BingOfficesUnsupportedFeatureException>(() => exporter.Export(request, output));
        Assert.Equal(BingOfficesStage.Preflight, error.Stage);
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, error.Code);
        Assert.Equal("NPOI", error.Provider);
        Assert.Equal(original, template.ToArray());
        Assert.Equal(previousOutput, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 核对全部来源单元格、透视字段和逐值缓存。
    /// </summary>
    /// <param name="bytes">实际公开导出的工作簿内容。</param>
    /// <param name="advanced">是否使用高级多度量模式。</param>
    /// <param name="target">实际透视目标工作表名称。</param>
    /// <param name="sheetNames">期望的全部工作表顺序；为空时使用默认业务与汇总工作表。</param>
    internal static void AssertWorkbook(byte[] bytes, bool advanced, string target = "Summary",
        IReadOnlyList<string> sheetNames = null)
    {
        using var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        Assert.Equal(sheetNames ?? new[] { "Data", "Summary" }, Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
        var data = workbook.GetSheet("Data");
        var headers = new[] { "Region", "Amount", "State", "Channel", "Code" };
        Assert.Equal(headers, Enumerable.Range(0, 5).Select(column => data.GetRow(0).GetCell(column).StringCellValue));
        var rows = Rows();
        for (var row = 0; row < rows.Length; row++)
        {
            var actual = data.GetRow(row + 1);
            Assert.Equal(rows[row].Region, actual.GetCell(0).StringCellValue);
            Assert.Equal(rows[row].Amount, actual.GetCell(1).NumericCellValue);
            Assert.Equal(rows[row].State, actual.GetCell(2).StringCellValue);
            Assert.Equal(rows[row].Channel, actual.GetCell(3).StringCellValue);
            Assert.Equal(rows[row].Code, actual.GetCell(4).StringCellValue);
        }
        Assert.Single(((XSSFSheet)workbook.GetSheet(target)).GetPivotTables());
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cache = ReadXml(archive, "pivotCacheDefinition");
        var records = ReadXml(archive, "pivotCacheRecords");
        var pivot = ReadXml(archive, "pivotTable");
        var fields = cache.Root.Element(ns + "cacheFields").Elements(ns + "cacheField").ToArray();
        Assert.Equal(headers, fields.Select(field => (string)field.Attribute("name")));
        Assert.Equal("A1:E4", (string)cache.Root.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref"));
        Assert.Equal(3, (int)cache.Root.Attribute("recordCount"));
        var recordRows = records.Root.Elements(ns + "r").ToArray();
        Assert.Equal(rows.Length, recordRows.Length);
        for (var row = 0; row < rows.Length; row++)
        {
            var indexes = recordRows[row].Elements(ns + "x").Select(item => (int)item.Attribute("v")).ToArray();
            Assert.Equal(5, indexes.Length);
            var shared = fields.Select((field, column) => field.Element(ns + "sharedItems").Elements().ElementAt(indexes[column])).ToArray();
            Assert.Equal(rows[row].Region, (string)shared[0].Attribute("v"));
            Assert.Equal(rows[row].Amount, (double)shared[1].Attribute("v"));
            Assert.Equal(rows[row].State, (string)shared[2].Attribute("v"));
            Assert.Equal(rows[row].Channel, (string)shared[3].Attribute("v"));
            Assert.Equal(rows[row].Code, (string)shared[4].Attribute("v"));
        }
        var dataFields = pivot.Root.Element(ns + "dataFields").Elements(ns + "dataField").ToArray();
        Assert.Equal(advanced ? 2 : 1, dataFields.Length);
        Assert.All(dataFields, field =>
        {
            Assert.Equal("-1", (string)field.Attribute("baseField"));
            Assert.Equal("1048832", (string)field.Attribute("baseItem"));
        });
        if (advanced)
        {
            Assert.Equal(new[] { "sum", "count" }, dataFields
                .Select(field => (string)field.Attribute("subtotal") ?? "sum"));
            Assert.Equal(new[] { 3, -2 }, pivot.Root.Element(ns + "colFields").Elements().Select(field => (int)field.Attribute("x")));
            Assert.Single(pivot.Root.Element(ns + "pageFields").Elements());
        }
    }

    /// <summary>
    /// 读取指定类型的唯一原生 XML 部件。
    /// </summary>
    /// <param name="archive">实际输出的工作簿包。</param>
    /// <param name="prefix">部件文件名类型前缀。</param>
    /// <returns>完整解析后的原生部件。</returns>
    private static XDocument ReadXml(ZipArchive archive, string prefix)
    {
        var entry = archive.Entries.Single(item => Path.GetFileName(item.FullName).StartsWith(prefix, StringComparison.Ordinal)
            && item.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 创建可重复使用的采购透视请求。
    /// </summary>
    /// <param name="advanced">是否使用高级多度量模式。</param>
    /// <param name="endRow">来源末行的零基坐标。</param>
    /// <returns>公开构建器产生的不可变请求。</returns>
    internal static ExcelWorkbookExportRequest Request(bool advanced, int endRow = 3) =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", Rows()).PivotTable(Definition(advanced, endRow)));

    /// <summary>
    /// 创建基础或高级采购透视定义。
    /// </summary>
    /// <param name="advanced">是否使用高级多度量模式。</param>
    /// <param name="endRow">来源末行的零基坐标。</param>
    /// <param name="target">新建透视目标工作表。</param>
    /// <returns>包含全部来源字段的定义。</returns>
    private static ExcelPivotTableDefinition Definition(bool advanced, int endRow, string target = "Summary") => new()
    {
        Name = "PurchasePivot", SourceSheet = "Data", TargetSheet = target,
        SourceRange = new ExcelRangeDefinition { StartRow = 0, EndRow = endRow, StartColumn = 0, EndColumn = 4 },
        RowField = advanced ? null : "Region", ValueField = advanced ? null : "Amount",
        RowFields = advanced ? new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } } : Array.Empty<ExcelPivotAxisFieldDefinition>(),
        ColumnFields = advanced ? new[] { new ExcelPivotAxisFieldDefinition { Field = "Channel" } } : Array.Empty<ExcelPivotAxisFieldDefinition>(),
        ValueFields = advanced ? new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" },
            new ExcelPivotValueDefinition { Field = "Code", Caption = "CodeCount", Aggregation = ExcelPivotAggregation.Count }
        } : Array.Empty<ExcelPivotValueDefinition>(),
        Filters = advanced ? new[] { new ExcelPivotFilterDefinition { Field = "State", Values = new object[] { "Open" } } } : Array.Empty<ExcelPivotFilterDefinition>()
    };

    /// <summary>
    /// 创建完整采购来源记录。
    /// </summary>
    /// <returns>包含筛选外记录的业务来源。</returns>
    private static NpoiPivotExportRow[] Rows() => new[]
    {
        new NpoiPivotExportRow { Region = "North", Amount = 10, State = "Open", Channel = "Retail", Code = "A" },
        new NpoiPivotExportRow { Region = "North", Amount = 20, State = "Closed", Channel = "Wholesale", Code = "B" },
        new NpoiPivotExportRow { Region = "South", Amount = 30, State = "Open", Channel = "Retail", Code = "C" }
    };
}
