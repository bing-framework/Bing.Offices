using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using NPOI.SS.UserModel;
using Xunit;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 图片和原生校验的独立格式合同。
/// </summary>
public sealed class SheetContentContractTest
{
    // 预期独立维护，不能由生产能力声明推导。
    /// <summary>
    /// 独立维护的提供程序与格式的图片及校验支持预期。
    /// </summary>
    private static readonly Dictionary<string, bool> Profiles = new()
    {
        ["NPOI/Xlsx"] = true, ["NPOI/Xls"] = true,
        ["ClosedXML/Xlsx"] = true, ["MiniExcel/Xlsx"] = false, ["SpreadCheetah/Xlsx"] = true
    };
    /// <summary>
    /// 用于验证原始图片字节保留的单像素 PNG 样本。
    /// </summary>
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=");
    /// <summary>
    /// 提供图片和校验合同的格式及调用方式组合。
    /// </summary>
    /// <returns>提供程序、工作簿格式和同步或异步入口组合。</returns>
    public static IEnumerable<object[]> Cases()
    {
        foreach (var driver in ProviderDrivers.All)
        foreach (var async in new[] { false, true })
            yield return new object[] { driver.Name, ExcelFormat.Xlsx, async };
        foreach (var async in new[] { false, true })
        {
            yield return new object[] { "NPOI", ExcelFormat.Xls, async };
            yield return new object[] { "SpreadCheetah", ExcelFormat.Xlsx, async };
        }
    }

    /// <summary>
    /// 验证日期校验保留 1900 年初日期序列值。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="month">1900 年测试日期的月份。</param>
    /// <param name="day">1900 年测试日期的日。</param>
    /// <param name="serial">预期的 Excel 日期序列值。</param>
    [Theory]
    [InlineData("NPOI", 1, 1, 1)]
    [InlineData("NPOI", 2, 28, 59)]
    [InlineData("ClosedXML", 1, 1, 1)]
    [InlineData("ClosedXML", 2, 28, 59)]
    [InlineData("SpreadCheetah", 1, 1, 1)]
    [InlineData("SpreadCheetah", 2, 28, 59)]
    public async Task Export_DateValidation_ShouldPreserveEarly1900Serial(string provider, int month, int day, int serial)
    {
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows", new[] { new Row { Name = "日期" } },
            sheet => sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Type = ExcelDataValidationType.Date, Operator = ExcelConditionalComparisonOperator.Equal,
                Date1 = new DateTime(1900, month, day), Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1 }
            })));
        using var output = new MemoryStream();
        await Export(provider, request, output, true);
        using var package = new ZipArchive(new MemoryStream(output.ToArray()));
        using var input = package.GetEntry("xl/worksheets/sheet1.xml").Open();
        var xml = XDocument.Load(input);
        Assert.Equal(serial.ToString(System.Globalization.CultureInfo.InvariantCulture),
            xml.Descendants().Single(element => element.Name.LocalName == "formula1").Value.TrimStart('='));
    }

    /// <summary>
    /// 提供报表地址和冻结窗格越界用例。
    /// </summary>
    /// <returns>提供程序、格式、调用方式和无效边界场景组合。</returns>
    public static IEnumerable<object[]> ReportBoundaryCases()
    {
        foreach (var provider in new[] { "NPOI", "ClosedXML" })
        foreach (var format in provider == "NPOI" ? new[] { ExcelFormat.Xls, ExcelFormat.Xlsx } : new[] { ExcelFormat.Xlsx })
        foreach (var async in new[] { false, true })
        foreach (var scenario in new[] { "reverse-name", "reverse-print", "print-row-overflow", "repeat-rows", "repeat-columns", "freeze-origin" })
            yield return new object[] { provider, format, async, scenario };
    }

    /// <summary>
    /// 验证条件格式的默认色阶和数据条颜色一致。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Export_ConditionalDefaults_ShouldUseEquivalentColors(string provider, bool async)
    {
        var range = new ExcelRangeDefinition { StartRow = 1, EndRow = 2 };
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows", new[] { new Row { Name = "完整内容" } },
            sheet => sheet.ConditionalFormat(new ExcelConditionalFormatDefinition { Type = ExcelConditionalFormatType.ColorScale, Range = range })
                .ConditionalFormat(new ExcelConditionalFormatDefinition { Type = ExcelConditionalFormatType.DataBar, Range = range })));
        using var output = new MemoryStream();
        await Export(provider, request, output, async);
        using var package = new ZipArchive(new MemoryStream(output.ToArray()));
        using var input = package.GetEntry("xl/worksheets/sheet1.xml").Open();
        var xml = XDocument.Load(input);
        var scale = xml.Descendants().Single(element => element.Name.LocalName == "colorScale");
        var colors = scale.Elements().Where(element => element.Name.LocalName == "color")
            .Select(element => ((string)element.Attribute("rgb"))[^6..]).ToArray();
        Assert.Equal(new[] { "FF0000", "00FF00" }, colors);
        var bar = xml.Descendants(xml.Root.Name.Namespace + "dataBar").Single();
        Assert.EndsWith("0000FF", (string)bar.Elements().Single(element => element.Name.LocalName == "color").Attribute("rgb"));
    }

    /// <summary>
    /// 验证无效报表边界在枚举和输出前被拒绝。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="format">待验证的工作簿格式。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    /// <param name="scenario">独立能力预期对应的合同场景。</param>
    [Theory]
    [MemberData(nameof(ReportBoundaryCases))]
    public async Task Export_ReportBoundary_ShouldRejectBeforeEnumeration(string provider, ExcelFormat format, bool async, string scenario)
    {
        var enumerations = 0;
        IEnumerable<Row> Rows()
        {
            enumerations++;
            yield return new Row { Name = "不得读取" };
        }
        var request = ExcelExport.Workbook(book => book.Format(format).AddSheet("Rows", Rows(), sheet =>
        {
            switch (scenario)
            {
                case "reverse-name": sheet.NamedRange(new ExcelNamedRangeDefinition { Name = "Values", Address = "B2:A1" }); break;
                case "reverse-print": sheet.PrintLayout(new ExcelPrintLayoutOptions { PrintArea = "B2:A1" }); break;
                case "print-row-overflow": sheet.PrintLayout(new ExcelPrintLayoutOptions { PrintArea = "A1:A999999999999999" }); break;
                case "repeat-rows": sheet.PrintLayout(new ExcelPrintLayoutOptions { RepeatRows = "3:2" }); break;
                case "repeat-columns": sheet.PrintLayout(new ExcelPrintLayoutOptions { RepeatColumns = "XFE:XFE" }); break;
                default: sheet.FreezePane(new ExcelFreezePaneDefinition { Rows = 1, TopRow = 1048576 }); break;
            }
        }));
        using var output = new MemoryStream();
        var error = await Record.ExceptionAsync(() => Export(provider, request, output, async));
        var structured = Assert.IsAssignableFrom<BingOfficesException>(error);
        Assert.Equal(BingOfficesStage.Preflight, structured.Stage);
        Assert.Equal(0, enumerations);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证图片内容及各类数据校验运算符完整写入。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="format">待验证的工作簿格式。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task Export_ShouldPreserveImageAndEveryValidationOperator(string provider, ExcelFormat format, bool async)
    {
        Assert.True(Profiles.TryGetValue(provider + "/" + format, out var supported), "Missing independent content profile");
        var definitions = new List<ExcelDataValidationDefinition>();
        foreach (var type in new[] { ExcelDataValidationType.Integer, ExcelDataValidationType.Decimal, ExcelDataValidationType.Date, ExcelDataValidationType.TextLength })
        foreach (ExcelConditionalComparisonOperator op in Enum.GetValues(typeof(ExcelConditionalComparisonOperator)))
            definitions.Add(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition { StartRow = definitions.Count + 1, EndRow = definitions.Count + 1, EndColumn = 0 },
                Type = type, Operator = op, Value1 = 1, Value2 = 9,
                Date1 = new DateTime(2024, 1, 1), Date2 = new DateTime(2025, 1, 1),
                IgnoreBlanks = definitions.Count % 2 == 0, InputTitle = "输入", InputMessage = "请输入合法值", ErrorTitle = "错误", ErrorMessage = "值无效"
            });
        definitions.Add(new ExcelDataValidationDefinition
        {
            Type = ExcelDataValidationType.List, Range = new ExcelRangeDefinition { StartRow = 34, EndRow = 35 },
            Values = new[] { "中文", "第二项" }
        });
        var request = ExcelExport.Workbook(book => book.Format(format).AddSheet("Rows", new[] { new Row { Name = "完整内容" } }, sheet =>
        {
            sheet.Image(new ExcelSheetImageDefinition { Content = Png, Row = 2, Column = 2, Width = 40, Height = 30, OffsetX = 3, OffsetY = 4 });
            foreach (var definition in definitions) sheet.DataValidation(definition);
        }));
        using var output = new MemoryStream();
        if (!supported)
        {
            var error = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() => Export(provider, request, output, async));
            Assert.Equal(BingOfficesStage.Preflight, error.Stage);
            Assert.Equal(provider, error.Provider);
            Assert.Equal(BingOfficesOperation.Export, error.Operation);
            Assert.Empty(output.ToArray());
            return;
        }
        await Export(provider, request, output, async);
        Assert.True(output.CanWrite);
        var bytes = output.ToArray();
        if (format == ExcelFormat.Xlsx)
        {
            using var package = new ZipArchive(new MemoryStream(bytes));
            using var input = package.GetEntry("xl/worksheets/sheet1.xml").Open();
            var xml = XDocument.Load(input);
            Assert.Equal("\"中文,第二项\"", xml.Descendants().Last(element => element.Name.LocalName == "formula1").Value);
        }
        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes));
        Assert.Equal(1, workbook.NumberOfSheets);
        var sheetResult = workbook.GetSheetAt(0);
        Assert.Equal("Rows", sheetResult.SheetName);
        Assert.Equal("Name", sheetResult.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("完整内容", sheetResult.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal(Png, Assert.Single(workbook.GetAllPictures().Cast<IPictureData>()).Data);
        var validations = sheetResult.GetDataValidations();
        Assert.Equal(definitions.Count, validations.Count);
        for (var index = 0; index < definitions.Count; index++)
        {
            var expected = definitions[index];
            var actual = validations[index];
            Assert.Equal(expected.IgnoreBlanks, actual.EmptyCellAllowed);
            Assert.Equal(expected.Range.StartRow, actual.Regions.CellRangeAddresses[0].FirstRow);
            Assert.Equal(expected.Range.EndRow, actual.Regions.CellRangeAddresses[0].LastRow);
            if (expected.Type == ExcelDataValidationType.List)
                Assert.Equal("\"" + string.Join(",", expected.Values) + "\"", actual.ValidationConstraint.Formula1
                    ?? "\"" + string.Join(",", actual.ValidationConstraint.ExplicitListValues ?? Array.Empty<string>()) + "\"");
            else
            {
                Assert.Equal(ExpectedOperator(expected.Operator), actual.ValidationConstraint.Operator);
                Assert.Equal(expected.ErrorTitle, actual.ErrorBoxTitle);
                Assert.Equal(expected.ErrorMessage, actual.ErrorBoxText);
                Assert.Equal(expected.InputTitle, actual.PromptBoxTitle);
                Assert.Equal(expected.InputMessage, actual.PromptBoxText);
                Assert.True(actual.ShowErrorBox);
                Assert.True(actual.ShowPromptBox);
                var expectedType = expected.Type switch { ExcelDataValidationType.Integer => 1, ExcelDataValidationType.Decimal => 2, ExcelDataValidationType.Date => 4, _ => 6 };
                Assert.Equal(expectedType, actual.ValidationConstraint.GetValidationType());
                Assert.Equal(expected.Type == ExcelDataValidationType.Date ? "45292" : "1", actual.ValidationConstraint.Formula1
                    ?? (actual.ValidationConstraint as NPOI.HSSF.UserModel.DVConstraint)?.Value1.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
        }
        if (format == ExcelFormat.Xlsx)
        {
            using var zip = new ZipArchive(new MemoryStream(bytes));
            var drawing = Assert.Single(zip.Entries.Where(entry => entry.FullName.StartsWith("xl/drawings/drawing", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)));
            using var drawingInput = drawing.Open();
            var xml = XDocument.Load(drawingInput);
            Assert.Single(xml.Descendants().Where(element => element.Name.LocalName == "pic"));
            Assert.Contains(zip.Entries, entry => entry.FullName.StartsWith("xl/media/", StringComparison.Ordinal));
            Assert.Contains(zip.Entries, entry => entry.FullName.StartsWith("xl/drawings/_rels/", StringComparison.Ordinal));
            var from = Assert.Single(xml.Descendants().Where(element => element.Name.LocalName == "from"));
            Assert.Equal("2", from.Elements().Single(element => element.Name.LocalName == "row").Value);
            Assert.Equal("2", from.Elements().Single(element => element.Name.LocalName == "col").Value);
            Assert.Equal("28575", from.Elements().Single(element => element.Name.LocalName == "colOff").Value);
            Assert.Equal("38100", from.Elements().Single(element => element.Name.LocalName == "rowOff").Value);
        }
    }

    /// <summary>
    /// 验证日期校验遵循模板的 1904 日期系统。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [InlineData("NPOI", false)]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", false)]
    [InlineData("ClosedXML", true)]
    public async Task Export_DateValidation_ShouldRespect1904Template(string provider, bool async)
    {
        using var template = new MemoryStream();
        using (var workbook = new ClosedXML.Excel.XLWorkbook())
        {
            workbook.Use1904DateSystem = true;
            workbook.AddWorksheet("Rows");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .AddSheet("Rows", new[] { new Row { Name = "日期" } }, sheet => sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Type = ExcelDataValidationType.Date, Operator = ExcelConditionalComparisonOperator.Between,
                Date1 = new DateTime(1904, 1, 1), Date2 = new DateTime(1904, 1, 2),
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1 }
            })));
        using var output = new MemoryStream();
        await Export(provider, request, output, async);
        Assert.True(template.CanRead);
        using var package = new ZipArchive(new MemoryStream(output.ToArray()));
        using var input = package.GetEntry("xl/worksheets/sheet1.xml").Open();
        var xml = XDocument.Load(input);
        Assert.Equal("0", xml.Descendants().Single(element => element.Name.LocalName == "formula1").Value.TrimStart('='));
        Assert.Equal("1", xml.Descendants().Single(element => element.Name.LocalName == "formula2").Value.TrimStart('='));
    }

    /// <summary>
    /// 验证构建器复制图片及列表数据并拒绝无效定义。
    /// </summary>
    [Fact]
    public void Builder_ShouldSnapshotImageAndListAndRejectInvalidDefinitions()
    {
        var bytes = Png.ToArray();
        var values = new[] { "A", "B" };
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows", Array.Empty<Row>(), sheet => sheet
            .Image(new ExcelSheetImageDefinition { Content = bytes, Width = 1, Height = 1 })
            .DataValidation(new ExcelDataValidationDefinition { Type = ExcelDataValidationType.List, Values = values, Range = new ExcelRangeDefinition() })));
        bytes[0] = 0;
        values[0] = "changed";
        Assert.Equal(Png, request.Sheets[0].Images[0].Content);
        Assert.Equal(new[] { "A", "B" }, request.Sheets[0].DataValidations[0].Values);
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition { Range = new ExcelRangeDefinition(), Value1 = double.NaN }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition { Range = new ExcelRangeDefinition(), Type = ExcelDataValidationType.List, Values = new[] { "A,B" } }.Validate());
    }

    /// <summary>
    /// 验证跨工作表导出保留 JPEG 原始内容。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="format">待验证的工作簿格式。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task Export_JpegAcrossSheets_ShouldPreserveContents(string provider, ExcelFormat format, bool async)
    {
        using var pixels = new Image<Rgb24>(2, 3);
        using var jpeg = new MemoryStream();
        pixels.SaveAsJpeg(jpeg);
        var content = jpeg.ToArray();
        var request = ExcelExport.Workbook(book =>
        {
            book.Format(format);
            foreach (var name in new[] { "First", "Second" })
                book.AddSheet(name, new[] { new Row { Name = name } }, sheet => sheet.Image(
                    new ExcelSheetImageDefinition { Content = content, Width = 20, Height = 30, Row = 1, Column = 1 }));
        });
        using var output = new MemoryStream();
        if (!Profiles[provider + "/" + format])
        {
            await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() => Export(provider, request, output, async));
            Assert.Equal(0, output.Length);
            return;
        }
        await Export(provider, request, output, async);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray()));
        Assert.Equal(2, workbook.NumberOfSheets);
        Assert.All(workbook.GetAllPictures().Cast<IPictureData>(), picture => Assert.Equal(content, picture.Data));
        Assert.NotEmpty(workbook.GetAllPictures());
        for (var index = 0; index < 2; index++)
            Assert.Equal(index == 0 ? "First" : "Second", workbook.GetSheetAt(index).GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 验证超限校验区域在枚举和输出前被拒绝。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="format">待验证的工作簿格式。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [MemberData(nameof(Cases))]
    public async Task Export_OutOfRange_ShouldFailBeforeEnumerationOrOutput(string provider, ExcelFormat format, bool async)
    {
        var enumerated = false;
        IEnumerable<Row> Rows()
        {
            enumerated = true;
            yield return new Row();
        }
        var request = ExcelExport.Workbook(book => book.Format(format).AddSheet("Rows", Rows(), sheet => sheet.DataValidation(
            new ExcelDataValidationDefinition { Range = new ExcelRangeDefinition { EndRow = format == ExcelFormat.Xls ? 65536 : 1048576 } })));
        using var output = new MemoryStream();
        var error = await Record.ExceptionAsync(() => Export(provider, request, output, async));
        Assert.NotNull(error);
        Assert.False(enumerated);
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证图片导出枚举失败清理暂存文件并保留原目标。
    /// </summary>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SpreadImages_FailedEnumeration_ShouldCleanStagingAndPreserveFile(bool async)
    {
        var before = Directory.GetFiles(Path.GetTempPath(), "bing-spread-images-*.tmp");
        var directory = Path.Combine(Path.GetTempPath(), "sheet-content-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "output.xlsx");
        var sentinel = new byte[] { 1, 2, 3 };
        await File.WriteAllBytesAsync(path, sentinel);
        var expected = new InvalidOperationException("enumeration failed");
        IEnumerable<Row> Rows()
        {
            yield return new Row { Name = "first" };
            throw expected;
        }
        try
        {
            var request = ExcelExport.Workbook(book => book.AddSheet("Rows", Rows(), sheet => sheet.Image(
                new ExcelSheetImageDefinition { Content = Png, Width = 20, Height = 30 })));
            var exporter = new SpreadCheetahStreamingExcelExporter();
            var error = await Record.ExceptionAsync(async () =>
            {
                if (async) await exporter.ExportBatchesToFileAsync(request, path);
                else exporter.ExportBatchesToFile(request, path);
            });
            Assert.NotNull(error);
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
            Assert.Equal(before.OrderBy(value => value), Directory.GetFiles(Path.GetTempPath(), "bing-spread-images-*.tmp").OrderBy(value => value));
        }
        finally { Directory.Delete(directory, true); }
    }

    /// <summary>
    /// 获取公共比较运算符对应的原生校验编号。
    /// </summary>
    /// <param name="op">公共数据校验比较运算符。</param>
    /// <returns>原生数据校验使用的比较运算符编号。</returns>
    private static int ExpectedOperator(ExcelConditionalComparisonOperator op) => op switch
    {
        ExcelConditionalComparisonOperator.Between => 0, ExcelConditionalComparisonOperator.NotBetween => 1,
        ExcelConditionalComparisonOperator.Equal => 2, ExcelConditionalComparisonOperator.NotEqual => 3,
        ExcelConditionalComparisonOperator.GreaterThan => 4, ExcelConditionalComparisonOperator.LessThan => 5,
        ExcelConditionalComparisonOperator.GreaterThanOrEqual => 6, _ => 7
    };

    /// <summary>
    /// 通过指定提供程序执行同步或异步导出。
    /// </summary>
    /// <param name="provider">待验证的提供程序名称。</param>
    /// <param name="request">待执行的工作簿导出请求。</param>
    /// <param name="output">接收导出内容的目标流。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    private static async Task Export(string provider, ExcelWorkbookExportRequest request, Stream output, bool async)
    {
        if (provider == "SpreadCheetah")
        {
            if (async) await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, output);
            else new SpreadCheetahStreamingExcelExporter().ExportBatches(request, output);
            return;
        }
        var exporter = ProviderDrivers.Get(provider).CreateExporter();
        if (async) await exporter.ExportAsync(request, output);
        else exporter.Export(request, output);
    }

    /// <summary>
    /// 用于图片和数据校验合同的数据行。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; }
    }
}
