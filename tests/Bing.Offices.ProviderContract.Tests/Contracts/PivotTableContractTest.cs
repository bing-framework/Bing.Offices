using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证公共透视表定义在 Provider 工作簿中的结构和预检行为。
/// </summary>
public sealed class PivotTableContractTest
{
    /// <summary>
    /// 验证 Provider 写出完整透视表及其数据缓存。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider。</param>
    [Theory]
    [InlineData("ClosedXML")]
    [InlineData("NPOI")]
    public void XlsxExport_ShouldWritePivotTableAndCompleteCache(string provider)
    {
        Assert.True(provider == "NPOI"
            ? new NpoiExcelExporter().Supports(ExcelProviderFeatures.PivotTables)
            : new ClosedXmlExcelExporter().Supports(ExcelProviderFeatures.PivotTables));
        var range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1 };
        var definition = CreateDefinition(range);
        var request = CreateRequest(definition);

        Assert.NotSame(definition, request.PivotTables.Single());
        Assert.NotSame(range, request.PivotTables.Single().SourceRange);

        using var output = new MemoryStream();
        Export(provider, request, output);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var pivot = ReadXml(archive, archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName);
        var cache = ReadXml(archive, archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheDefinition1.xml", StringComparison.Ordinal)).FullName);
        var records = ReadXml(archive, archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheRecords1.xml", StringComparison.Ordinal)).FullName);
        var workbook = ReadXml(archive, "xl/workbook.xml");

        Assert.Equal("SalesByRegion", (string)pivot.Root.Attribute("name"));
        var location = (string)pivot.Root.Element(SpreadsheetNs + "location").Attribute("ref");
        Assert.Equal("C2", location.Split(':')[0]);
        var rowFields = pivot.Root.Element(SpreadsheetNs + "rowFields")
            .Elements().ToArray();
        Assert.True(rowFields.Length == 1, pivot.ToString());
        Assert.Equal("0", (string)rowFields[0].Attribute("x"));
        var dataField = pivot.Root.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField").Single();
        Assert.Equal("1", (string)dataField.Attribute("fld"));
        Assert.True(dataField.Attribute("subtotal") == null
            || (string)dataField.Attribute("subtotal") == "sum");

        var source = cache.Root.Element(SpreadsheetNs + "cacheSource")
            .Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Data", (string)source.Attribute("sheet"));
        Assert.Equal("A1:B4", (string)source.Attribute("ref"));
        var fields = cache.Root.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").Select(field => (string)field.Attribute("name")).ToArray();
        Assert.Equal(new[] { "Region", "Amount" }, fields);
        Assert.True((string)records.Root.Attribute("count") == "3"
            || records.Root.Elements(SpreadsheetNs + "r").Count() == 3, records.ToString());
        Assert.True(records.Root.Elements(SpreadsheetNs + "r").Count() == 3, records.ToString());
        var amountItems = cache.Root.Element(SpreadsheetNs + "cacheFields").Elements().ElementAt(1)
            .Element(SpreadsheetNs + "sharedItems").Elements().ToArray();
        Assert.Equal(new[] { "10", "20", "30" }, records.Root.Elements(SpreadsheetNs + "r")
            .Select(row => row.Elements().Last()).Select(item => item.Name.LocalName == "x"
                ? (string)amountItems[int.Parse((string)item.Attribute("v"))].Attribute("v")
                : (string)item.Attribute("v")).ToArray());
        var rowItems = cache.Root.Element(SpreadsheetNs + "cacheFields").Elements().First()
            .Element(SpreadsheetNs + "sharedItems").Elements().Select(item => (string)item.Attribute("v"))
            .ToArray();
        Assert.Equal(new[] { "North", "South", "North" }, records.Root.Elements(SpreadsheetNs + "r")
            .Select(row => row.Elements().First()).Select(item => item.Name.LocalName == "x"
                ? rowItems[int.Parse((string)item.Attribute("v"))] : (string)item.Attribute("v")).ToArray());
        using var reopened = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "Summary" }, reopened.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "Region", "Amount", "North", "10", "South", "20", "North", "30" },
            reopened.Worksheet("Data").Range("A1:B4").Cells().Select(cell => cell.GetString()));
        Assert.Single(workbook.Root.Element(SpreadsheetNs + "pivotCaches")
            .Elements(SpreadsheetNs + "pivotCache"));
    }

    /// <summary>
    /// 验证 NPOI 在 XLS 导出前拒绝透视请求。
    /// </summary>
    /// <param name="format">请求格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    public void Npoi_ShouldRejectBeforeWritingDestination(ExcelFormat format)
    {
        Assert.True(new NpoiExcelExporter().Supports(ExcelProviderFeatures.PivotTables));
        var definition = CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        });
        var request = ExcelExport.Workbook(builder => builder
            .Format(format)
            .AddSheet("Data", CreateRows())
            .PivotTable(definition));
        using var output = SeededStream();
        var original = output.ToArray();

        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(request, output));

        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证无效字段和重复目标 Sheet 在输出时失败且保持目标流不变。
    /// </summary>
    [Theory]
    [InlineData("ClosedXML")]
    public void InvalidField_ShouldFailWithoutChangingDestination(string provider)
    {
        var definition = CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        });
        definition = new ExcelPivotTableDefinition
        {
            Name = definition.Name,
            SourceSheet = definition.SourceSheet,
            SourceRange = definition.SourceRange,
            TargetSheet = definition.TargetSheet,
            TargetRow = definition.TargetRow,
            TargetColumn = definition.TargetColumn,
            RowField = "Missing",
            ValueField = definition.ValueField
        };
        var request = CreateRequest(definition);
        using var output = SeededStream();
        var original = output.ToArray();

        Assert.Throws<BingOfficesConfigurationException>(() => Export(provider, request, output));

        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证非数值求和列在写入前被拒绝。
    /// </summary>
    [Fact]
    public void NonNumericValueField_ShouldFailWithoutChangingDestination()
    {
        var definition = CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        });
        var request = CreateRequest(new ExcelPivotTableDefinition
        {
            Name = definition.Name,
            SourceSheet = definition.SourceSheet,
            SourceRange = definition.SourceRange,
            TargetSheet = definition.TargetSheet,
            TargetRow = definition.TargetRow,
            TargetColumn = definition.TargetColumn,
            RowField = "Amount",
            ValueField = "Region"
        });
        using var output = SeededStream();
        var original = output.ToArray();

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证前向 Provider 明确拒绝透视表请求。
    /// </summary>
    /// <param name="provider">待验证的 Provider。</param>
    [Theory]
    [InlineData("MiniExcel")]
    [InlineData("SpreadCheetah")]
    public void StreamingProvider_ShouldRejectBeforeWritingDestination(string provider)
    {
        var request = CreateRequest(CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        }));
        using var output = SeededStream();
        var original = output.ToArray();

        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
        {
            if (provider == "MiniExcel")
                new MiniExcelExcelExporter().Export(request, output);
            else
                new SpreadCheetahStreamingExcelExporter().ExportBatches(request, output);
        });

        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证模板已占用透视目标工作表时保全调用方目标流。
    /// </summary>
    [Fact]
    public void TemplateTargetSheet_ShouldRejectWithoutChangingDestination()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Summary");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(builder => builder
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", CreateRows())
            .PivotTable(CreateDefinition(new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
            })));
        using var output = SeededStream();
        var original = output.ToArray();

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证同一工作簿中的透视表不能声明重复目标工作表。
    /// </summary>
    [Fact]
    public void DuplicateTargetSheet_ShouldBeRejectedWhenBuildingRequest()
    {
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder
            .AddSheet("Data", CreateRows())
            .PivotTable(CreateDefinition(new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
            }))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "OtherPivot",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
                },
                TargetSheet = "Summary",
                TargetRow = 5,
                TargetColumn = 0,
                RowField = "Region",
                ValueField = "Amount"
            })));
    }

    /// <summary>
    /// 验证预取消请求不会写入目标流。
    /// </summary>
    [Theory]
    [InlineData("ClosedXML")]
    public void PreCanceledExport_ShouldLeaveDestinationUnchanged(string provider)
    {
        var request = CreateRequest(CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        }));
        using var output = SeededStream();
        var original = output.ToArray();
        using var cancellation = new System.Threading.CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => Export(provider, request, output,
            cancellation.Token));

        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 验证透视目标不能成为后续透视表的业务来源。
    /// </summary>
    /// <param name="reverse">是否反转透视定义顺序。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PivotSource_ShouldRejectRegardlessOfDefinitionOrder(bool reverse)
    {
        var first = CreateDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        });
        var second = new ExcelPivotTableDefinition
        {
            Name = "Second", SourceSheet = "Summary", TargetSheet = "Second",
            SourceRange = first.SourceRange, RowField = "Region", ValueField = "Amount"
        };
        var enumerated = false;
        var exception = Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder =>
        {
            builder.AddSheet("Data", Rows());
            foreach (var definition in reverse ? new[] { second, first } : new[] { first, second })
                builder.PivotTable(definition);
        }));
        Assert.Contains("Summary", exception.Message);
        Assert.False(enumerated);

        System.Collections.Generic.IEnumerable<PivotRow> Rows()
        {
            enumerated = true;
            foreach (var row in CreateRows())
                yield return row;
        }
    }

    /// <summary>
    /// 验证多个透视表可独立引用同一业务工作表。
    /// </summary>
    [Fact]
    public void MultiplePivots_ShouldPreserveBusinessSourceAndAllDefinitions()
    {
        var range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1 };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", CreateRows())
            .PivotTable(CreateDefinition(range))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "Second", SourceSheet = "data", TargetSheet = "第二汇总",
                SourceRange = range, RowField = "region", ValueField = "amount"
            }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(new[] { "Data", "Summary", "第二汇总" }, workbook.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "SalesByRegion", "Second" }, workbook.Worksheets.SelectMany(sheet => sheet.PivotTables)
            .Select(pivot => pivot.Name));
        Assert.Equal(new[] { "Region", "Amount", "North", "10", "South", "20", "North", "30" },
            workbook.Worksheet("Data").Range("A1:B4").Cells().Select(cell => cell.GetString()));
    }

    /// <summary>
    /// 验证非法目标名称在枚举业务数据前被拒绝。
    /// </summary>
    /// <param name="target">目标工作表名称。</param>
    [Theory]
    [InlineData("Bad/Sheet")]
    [InlineData("Bad:Sheet")]
    [InlineData("Bad\\Sheet")]
    [InlineData("Bad?Sheet")]
    [InlineData("Bad*Sheet")]
    [InlineData("Bad[Sheet]")]
    [InlineData("12345678901234567890123456789012")]
    [InlineData("'Sheet")]
    [InlineData("Sheet'")]
    [InlineData("Bad\0Sheet")]
    [InlineData("Bad\u0003Sheet")]
    public void InvalidTargetName_ShouldFailBeforeEnumeratingBusinessRows(string target)
    {
        var enumerated = false;
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder.AddSheet("Data", Rows())
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "Pivot", SourceSheet = "Data", TargetSheet = target,
                SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1 },
                RowField = "Region", ValueField = "Amount"
            })));
        Assert.False(enumerated);

        System.Collections.Generic.IEnumerable<PivotRow> Rows()
        {
            enumerated = true;
            yield return new PivotRow { Region = "North", Amount = 10 };
        }
    }

    /// <summary>
    /// 验证透视展开区域不能超出 XLSX 边界。
    /// </summary>
    /// <param name="row">零基目标行。</param>
    /// <param name="column">零基目标列。</param>
    [Theory]
    [InlineData(1048573, 0)]
    [InlineData(1048575, 0)]
    [InlineData(0, 16383)]
    [InlineData(int.MaxValue, int.MaxValue)]
    public void PivotFootprint_ShouldRejectOverflowWithoutChangingDestination(int row, int column)
    {
        var request = CreateRequest(new ExcelPivotTableDefinition
        {
            Name = "Pivot", SourceSheet = "Data", TargetSheet = "Summary", TargetRow = row, TargetColumn = column,
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1 },
            RowField = "Region", ValueField = "Amount"
        });
        using var output = SeededStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(new byte[] { 7, 11, 19, 23 }, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证表头、分组及总计恰好容纳时允许导出。
    /// </summary>
    /// <param name="includeBlank">是否包含空分组键。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PivotFootprint_ShouldAllowExactBoundaryWithDuplicateAndBlankGroups(bool includeBlank)
    {
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new PivotRow { Region = includeBlank ? "" : "South", Amount = 20 },
            new PivotRow { Region = "North", Amount = 10 },
            new PivotRow { Region = "NORTH", Amount = 30 }
        }).PivotTable(new ExcelPivotTableDefinition
        {
            Name = "Pivot", SourceSheet = "Data", TargetSheet = "1234567890123456789012345678901",
            TargetRow = 1048572, TargetColumn = 16382,
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1 },
            RowField = "Region", ValueField = "Amount"
        }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()));
        var pivot = ReadXml(archive, archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName);
        Assert.Equal("XFC1048573", (string)pivot.Root.Element(SpreadsheetNs + "location").Attribute("ref"));
        Assert.Equal("1", (string)pivot.Root.Attribute("rowGrandTotals") ?? "1");
        Assert.Equal("1", (string)pivot.Root.Attribute("colGrandTotals") ?? "1");
    }

    /// <summary>
    /// 创建透视工作簿请求。
    /// </summary>
    /// <param name="definition">透视表定义。</param>
    /// <returns>包含完整业务明细的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(ExcelPivotTableDefinition definition) =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", CreateRows()).PivotTable(definition));

    /// <summary>
    /// 创建区域求和透视定义。
    /// </summary>
    /// <param name="range">包含表头的来源区域。</param>
    /// <returns>区域求和透视定义。</returns>
    private static ExcelPivotTableDefinition CreateDefinition(ExcelRangeDefinition range) => new()
    {
        Name = "SalesByRegion",
        SourceSheet = "Data",
        SourceRange = range,
        TargetSheet = "Summary",
        TargetRow = 1,
        TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    };

    /// <summary>
    /// 创建包含重复分组的业务明细。
    /// </summary>
    /// <returns>按源顺序排列的全部明细。</returns>
    private static PivotRow[] CreateRows() => new PivotRow[]
    {
        new() { Region = "North", Amount = 10 },
        new() { Region = "South", Amount = 20 },
        new() { Region = "North", Amount = 30 }
    };

    /// <summary>
    /// 使用指定 Provider 导出透视工作簿。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">导出请求。</param>
    /// <param name="output">调用方输出流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void Export(string provider, ExcelWorkbookExportRequest request, Stream output,
        System.Threading.CancellationToken cancellationToken = default)
    {
        if (provider == "NPOI")
            new NpoiExcelExporter().Export(request, output, cancellationToken);
        else
            new ClosedXmlExcelExporter().Export(request, output, cancellationToken);
    }

    /// <summary>
    /// 创建带已有内容的目标流。
    /// </summary>
    /// <returns>位于起点的目标流。</returns>
    private static MemoryStream SeededStream()
    {
        var stream = new MemoryStream();
        stream.Write(new byte[] { 7, 11, 19, 23 });
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 读取工作簿 XML 部件。
    /// </summary>
    /// <param name="archive">工作簿包。</param>
    /// <param name="entryName">部件名称。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        Assert.True(entry != null, entryName + ": "
            + string.Join(", ", archive.Entries.Select(item => item.FullName)));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// SpreadsheetML 工作簿结构命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 区域销售明细。
    /// </summary>
    private sealed class PivotRow
    {
        /// <summary>
        /// 初始化一个 <see cref="PivotRow" /> 类型的实例。
        /// </summary>
        public PivotRow()
        {
        }

        /// <summary>
        /// 获取或设置区域。
        /// </summary>
        public string Region { get; set; }
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
