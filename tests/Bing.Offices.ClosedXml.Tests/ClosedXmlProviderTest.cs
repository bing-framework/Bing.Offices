using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML Provider 的直接职责测试；所有 Workbook 断言均通过真实 ClosedXML 读回。
/// </summary>
public sealed class ClosedXmlProviderTest
{
    /// <summary>
    /// 验证 ClosedXML 能力声明、DI 注册和首次注册优先语义。
    /// </summary>
    [Fact]
    public void CapabilitiesAndDi_ShouldBeProviderSpecificAndFirstRegistrationWins()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesClosedXml()
            .BuildServiceProvider();

        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        Assert.IsType<ClosedXmlExcelExporter>(exporter);
        Assert.IsType<ClosedXmlExcelImporter>(importer);
        Assert.Equal("ClosedXML", ((IExcelProviderCapabilities)exporter).ProviderName);
        Assert.True(((IExcelProviderCapabilities)exporter).Supports(
            ExcelProviderCapabilities.List | ExcelProviderCapabilities.Workbook
            | ExcelProviderCapabilities.Entity
            | ExcelProviderCapabilities.Xlsx));
        Assert.IsAssignableFrom<IExcelEntityExporter>(exporter);
        Assert.IsAssignableFrom<IExcelEntityImporter>(importer);
        Assert.IsType<ClosedXmlExcelExporter>(provider.GetRequiredService<IExcelEntityExporter>());
        Assert.IsType<ClosedXmlExcelImporter>(provider.GetRequiredService<IExcelEntityImporter>());

        using var first = new ServiceCollection()
            .AddBingOfficesClosedXml()
            .AddBingOfficesClosedXml()
            .BuildServiceProvider();
        Assert.IsType<ClosedXmlExcelExporter>(first.GetRequiredService<IExcelExporter>());

        using var configured = new ServiceCollection()
            .AddBingOfficesClosedXml(options =>
            {
                options.MaxConcurrentWorkbooks = 2;
                options.MaxQueuedOperations = 3;
            })
            .BuildServiceProvider();
        var configuredOptions = configured.GetRequiredService<ClosedXmlProviderOptions>();
        Assert.Equal(2, configuredOptions.MaxConcurrentWorkbooks);
        Assert.Equal(3, configuredOptions.MaxQueuedOperations);
    }

    /// <summary>
    /// 验证多工作表、样式和公式导出后可被 ClosedXML 读回。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteMultipleSheetsStylesAndFormula()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(new ExcelCellStyle { Bold = true })
                    .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00" }))
            .AddSheet("Formula", new[] { new FormulaRow { Formula = "=1+1" } }));
        using var stream = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, stream);
        Assert.True(stream.Length > 0);
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        Assert.Equal(new[] { "People", "Formula" }, workbook.Worksheets.Select(sheet => sheet.Name));
        var people = workbook.Worksheet("People");
        Assert.Equal("Name", people.Cell(1, 1).GetString());
        Assert.Equal("Alice", people.Cell(2, 1).GetString());
        Assert.Equal(30d, people.Cell(2, 2).GetDouble());
        Assert.True(people.Cell(1, 1).Style.Font.Bold);
        Assert.Equal("0.00", people.Cell(2, 3).Style.NumberFormat.Format);
        Assert.Equal("1+1", workbook.Worksheet("Formula").Cell(2, 1).FormulaA1);
    }

    /// <summary>
    /// 验证边框样式覆盖和模板对齐样式保留。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteBorderStyleAndPreserveTemplateAlignment()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell(1, 1).Value = "Template Header";
            sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var style = new ExcelCellStyle
        {
            Bold = true,
            TopBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor { Argb = "112233" }
            },
            BottomBorder = new ExcelBorderStyle { LineStyle = ExcelBorderLineStyle.Double }
        };
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(style)));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var header = workbook.Worksheet("People").Cell(1, 1);
        Assert.True(header.Style.Font.Bold);
        Assert.Equal(XLBorderStyleValues.Thin, header.Style.Border.TopBorder);
        Assert.Equal(XLBorderStyleValues.Double, header.Style.Border.BottomBorder);
        Assert.Equal(XLAlignmentHorizontalValues.Center, header.Style.Alignment.Horizontal);
    }

    /// <summary>
    /// 验证显式样式重置在模板样式叠加前生效。
    /// </summary>
    [Fact]
    public void Export_ShouldApplyExplicitStyleResetBeforeOverlay()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            var templateHeader = sheet.Cell(1, 1);
            templateHeader.Value = "Old";
            templateHeader.Style.Font.Bold = true;
            templateHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            templateHeader.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(new ExcelCellStyle
                {
                    Reset = new ExcelCellStyleReset
                    {
                        Bold = true,
                        HorizontalAlignment = true,
                        TopBorder = true
                    }
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var header = workbook.Worksheet("People").Cell(1, 1);
        Assert.False(header.Style.Font.Bold);
        Assert.Equal(XLAlignmentHorizontalValues.General, header.Style.Alignment.Horizontal);
        Assert.Equal(XLBorderStyleValues.None, header.Style.Border.TopBorder);
    }

    /// <summary>
    /// 验证自定义表头跨度和批注可以写入真实工作簿。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteCustomHeaderSpansAndComments()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2).HeaderRows(new[]
                {
                    new ExcelHeaderRow(0, new[]
                    {
                        new ExcelHeaderCell(0, "Profile", columnSpan: 2,
                            comment: new ExcelComment("generated", "tester")),
                        new ExcelHeaderCell(2, "Score")
                    })
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Profile", sheet.Cell("A1").GetString());
        Assert.Equal("Score", sheet.Cell("C1").GetString());
        Assert.Equal("A1:B1", Assert.Single(sheet.MergedRanges).RangeAddress.ToString());
        Assert.True(sheet.Cell("A1").HasComment);
        Assert.Equal("generated", sheet.Cell("A1").GetComment().ToString());
        Assert.Equal("tester", sheet.Cell("A1").GetComment().Author);
    }

    /// <summary>
    /// 验证映射、日期和数值字段能够完成导出导入往返。
    /// </summary>
    [Fact]
    public void RoundTrip_ShouldImportMappingAndDateNumberValues()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "Bob", Age = 41, Score = 12.5m } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var person = Assert.Single(result.Workbook.People);
        Assert.Equal("Bob", person.Name);
        Assert.Equal(41, person.Age);
        Assert.Equal(12.5m, person.Score);
        Assert.Equal(new[] { 1 }, result.Sheets.Single().SourceRows);
    }

    /// <summary>
    /// 验证最大行数资源限制在 ClosedXML 加载前报告结构化错误。
    /// </summary>
    [Fact]
    public void Import_MaxRows_ShouldRejectBeforeClosedXmlLoadWithStructuredResourceError()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[]
            {
                new Person { Name = "First", Age = 1, Score = 1m },
                new Person { Name = "Second", Age = 2, Score = 2m }
            }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Single(result.Sheets);
        Assert.Empty(result.Sheets[0].SourceRows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        Assert.Equal("People", error.SheetName);
        Assert.Equal(3, error.RowIndex);
    }

    /// <summary>
    /// 验证取消令牌在 ClosedXML 加载前阻止导入。
    /// </summary>
    [Fact]
    public void Import_Cancellation_ShouldStopBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "cancel" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().Import(source, request, cancellation.Token));
    }

    /// <summary>
    /// 验证异步非定位输入在超过输入上限时边读边拒绝。
    /// </summary>
    [Fact]
    public async Task ImportAsync_NonSeekableInputOverLimit_ShouldRejectDuringAsyncBuffering()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = new string('x', 512) } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        const int maximum = 32;
        using var source = new TrackingNonSeekableReadStream(bytes);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = maximum })
            .Sheet<Person>("People", root => root.People));

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().ImportAsync(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.InRange(source.BytesRead, maximum + 1, maximum + 1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证定位输入恰好达到输入字节上限时可以成功导入。
    /// </summary>
    [Fact]
    public async Task ImportAsync_SeekableInputExactlyAtMaxInputBytes_ShouldSucceed()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "exact-limit" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = generated.Length })
            .Sheet<Person>("People", root => root.People));

        generated.Position = 0;
        var result = await new ClosedXmlExcelImporter().ImportAsync(generated, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("exact-limit", Assert.Single(result.Workbook.People).Name);
        Assert.True(generated.CanRead);
    }

    /// <summary>
    /// 验证图片资源限制在 ClosedXML 加载前生效。
    /// </summary>
    [Fact]
    public void Import_PictureResourceLimit_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "picture" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.CreateEntry("xl/media/image1.png");
            using var image = entry.Open();
            image.Write(new byte[] { 1, 2, 3 }, 0, 3);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxPictureBytes = 2 })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Contains("图片", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证图片数量和总资源限制在 DOM 创建前生效。
    /// </summary>
    [Fact]
    public void Import_PictureCountAndTotalResourceLimits_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "picture" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var pair in new[]
            {
                (Name: "xl/media/image1.png", Bytes: new byte[] { 1, 2, 3 }),
                (Name: "xl/media/image2.png", Bytes: new byte[] { 4, 5, 6 })
            })
            {
                var entry = archive.CreateEntry(pair.Name);
                using var image = entry.Open();
                image.Write(pair.Bytes, 0, pair.Bytes.Length);
            }
        }
        var bytes = source.ToArray();

        using var countSource = new MemoryStream(bytes, writable: false);
        var countRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxPictures = 1 })
            .Sheet<Person>("People", root => root.People));
        var countException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(countSource, countRequest));

        using var totalSource = new MemoryStream(bytes, writable: false);
        var totalRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTotalPictureBytes = 5 })
            .Sheet<Person>("People", root => root.People));
        var totalException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(totalSource, totalRequest));

        Assert.Equal(BingOfficesStage.Preflight, countException.Stage);
        Assert.Equal(BingOfficesStage.Preflight, totalException.Stage);
        Assert.Contains("图片", countException.Message, StringComparison.Ordinal);
        Assert.Contains("图片", totalException.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证公式字符串属性保留导入的公式文本。
    /// </summary>
    [Fact]
    public void FormulaImport_ShouldPreserveFormulaTextForStringProperty()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Formula",
            new[] { new FormulaRow { Formula = "=1+1" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<FormulaWorkbook>(workbook =>
            workbook.Sheet<FormulaRow>("Formula", root => root.Rows));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("=1+1", Assert.Single(result.Workbook.Rows).Formula);
    }

    /// <summary>
    /// 验证隐藏行列、合并锚点、空白和错误单元格的导入边界。
    /// </summary>
    [Fact]
    public void ImportBoundaries_ShouldIncludeHiddenRowsColumnsAndReadMergeAnchorBlankAndError()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Boundary");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = "Empty";
            sheet.Cell("C1").Value = "Error";
            sheet.Cell("D1").Value = "Merged";
            sheet.Cell("A2").Value = "first";
            sheet.Cell("B2").Value = string.Empty;
            sheet.Cell("C2").Value = XLError.DivisionByZero;
            sheet.Cell("D2").Value = "anchor";
            sheet.Range("D2:E2").Merge();
            sheet.Cell("A3").Value = "hidden";
            sheet.Cell("B3").Value = string.Empty;
            sheet.Cell("C3").Value = XLError.NumberInvalid;
            sheet.Cell("D3").Value = "hidden-anchor";
            sheet.Row(3).Hide();
            sheet.Column(2).Hide();
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<BoundaryWorkbook>(workbook => workbook
            .Sheet<BoundaryRow>("Boundary", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "first", "hidden" }, result.Workbook.Rows.Select(row => row.Name));
        Assert.All(result.Workbook.Rows, row => Assert.Null(row.Empty));
        Assert.Equal("DivisionByZero", result.Workbook.Rows[0].Error);
        Assert.Equal("anchor", result.Workbook.Rows[0].Merged);
        Assert.Equal("hidden-anchor", result.Workbook.Rows[1].Merged);
    }

    /// <summary>
    /// 验证错误单元格转换到不兼容目标类型时报告转换错误。
    /// </summary>
    [Fact]
    public void ImportErrorCell_ShouldReportConversionForIncompatibleTargetType()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Errors");
            sheet.Cell("A1").Value = "Amount";
            sheet.Cell("A2").Value = XLError.DivisionByZero;
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook =>
            workbook.Sheet<ErrorNumberRow>("Errors", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ValueConversion
            && error.PropertyName == nameof(ErrorNumberRow.Amount));
        Assert.Empty(result.Workbook.Rows);
    }

    /// <summary>
    /// 验证 1904 日期系统的数值序列解析语义。
    /// </summary>
    [Fact]
    public void Date1904NumericSerial_ShouldUseWorkbookDateSystem()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Dates");
            sheet.Cell(1, 1).Value = "Date";
            sheet.Cell(2, 1).Value = 1d;
            workbook.SaveAs(source);
        }
        MarkWorkbookDate1904(source);
        using (var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var reader = new StreamReader(archive.GetEntry("xl/workbook.xml")!.Open(), Encoding.UTF8, true))
        {
            var workbookXml = reader.ReadToEnd();
            Assert.Contains("date1904=\"1\"", workbookXml);
        }
        Assert.True(ExcelXlsxZipPreflight.GetDate1904(source));
        source.Position = 0;
        var request = ExcelImport.Workbook<DateWorkbook>(workbook =>
            workbook.Sheet<DateRow>("Dates", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new DateTime(1904, 1, 2), Assert.Single(result.Workbook.Rows).Date);
    }

    /// <summary>
    /// 验证实体固定单元格、列表区域和合并区域的往返语义。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRoundTripFixedCellsListAndMerge()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "B2", item => item.Number)
            .Cell("Invoice", "B3", item => item.Customer)
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "B1", item => item.Title)
            .ListRegion("Lines", "A1", item => item.Lines));
        var source = new EntityInvoice
        {
            Number = 1001,
            Customer = "Customer A",
            Title = "Sales Order",
            Lines = new List<EntityLine>
            {
                new EntityLine { Code = "A", Quantity = 2 },
                new EntityLine { Code = "B", Quantity = 3 }
            }
        };
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var invoice = workbook.Worksheet("Invoice");
            Assert.Equal("Sales Order", invoice.Cell("A1").GetString());
            Assert.Equal(1001d, invoice.Cell("B2").GetDouble());
            Assert.Single(invoice.MergedRanges);
            var lines = workbook.Worksheet("Lines");
            Assert.Equal("Code", lines.Cell("A1").GetString());
            Assert.Equal("B", lines.Cell("A3").GetString());
        }

        using var input = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(1001, result.Entity.Number);
        Assert.Equal("Customer A", result.Entity.Customer);
        Assert.Equal("Sales Order", result.Entity.Title);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(line => line.Quantity));
    }

    /// <summary>
    /// 验证实体导入选项在 ClosedXML 加载工作簿前执行输入字节限制。
    /// </summary>
    [Fact]
    public void EntityImportOptions_MaxInputBytes_ShouldRejectBeforeClosedXmlLoad()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(
            new EntityInvoice { Title = "Resource limit" }, layout, generated);
        using var source = new MemoryStream(generated.ToArray());
        IExcelEntityResourceImporter importer = new ClosedXmlExcelImporter();
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxInputBytes = 1 });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntity(source, layout, options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体列表区域按 Order、PlacementKey 和 ColumnIndex 导出并导入动态字典列。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldPreserveHeadersValuesAndRoundTrip()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" },
                new() { PropertyName = nameof(EntityDynamicLine.Quantity), Title = "Qty" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "zone", Title = "Zone", DataTypeName = "string", Order = 0,
                    PlacementKey = $"before:{nameof(EntityDynamicLine.Quantity)}"
                },
                new() { Key = "early", Title = "Early", DataTypeName = "int32", Order = 10 },
                new() { Key = "late", Title = "Late", DataTypeName = "string", Order = 20 },
                new()
                {
                    Key = "indexed", Title = "Indexed", DataTypeName = "boolean", Order = 30,
                    ColumnIndex = 5
                }
            }
        };
        var layout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "B2", root => root.Lines, region => region.Mapping(mapping)));
        var source = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A", Quantity = 2,
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North", ["early"] = 7, ["late"] = "tail", ["indexed"] = true
                    }
                }
            }
        };
        using var exported = new MemoryStream();

        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("Dynamic");
            Assert.Equal(new[] { "Item", "Zone", "Qty", "Early", "Late", "Indexed" },
                sheet.Range("B2:G2").Cells().Select(cell => cell.GetString()));
            Assert.Equal("A", sheet.Cell("B3").GetString());
            Assert.Equal("North", sheet.Cell("C3").GetString());
            Assert.Equal(2d, sheet.Cell("D3").GetDouble());
            Assert.Equal(7d, sheet.Cell("E3").GetDouble());
            Assert.Equal("tail", sheet.Cell("F3").GetString());
            Assert.True(sheet.Cell("G3").GetBoolean());
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("A", line.Name);
        Assert.Equal(2, line.Quantity);
        Assert.Equal("North", line.CustomFields["zone"]);
        Assert.Equal(7, line.CustomFields["early"]);
        Assert.Equal("tail", line.CustomFields["late"]);
        Assert.Equal(true, line.CustomFields["indexed"]);
    }

    /// <summary>
    /// 验证实体动态列复用命名转换器和内置校验绑定。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldApplyConverterAndValidation()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "zone", Title = "Zone", DataTypeName = "string", ConverterName = "entity-title",
                    PlacementKey = $"before:{nameof(EntityDynamicLine.Quantity)}",
                    ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                    {
                        new() { Name = "maxLength", MaxLength = 20 }
                    }
                }
            }
        };
        var layout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines, region => region.Mapping(mapping)));
        var source = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        using var exported = new MemoryStream();
        var converters = new[] { new EntityTitleConverter() };
        new ClosedXmlExcelExporter(converters).ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("entity:North", workbook.Worksheet("Dynamic").Cell("B2").GetString());

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter(valueConverters: converters).ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("North", Assert.Single(result.Entity.Lines).CustomFields["zone"]);

        var invalid = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "B",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = new string('x', 21)
                    }
                }
            }
        };
        using var invalidOutput = new MemoryStream();
        Assert.Throws<BingOfficesExportException>(() =>
            new ClosedXmlExcelExporter(converters).ExportEntity(invalid, layout, invalidOutput));
        Assert.Empty(invalidOutput.ToArray());
    }

    /// <summary>
    /// 验证实体动态列与固定列冲突或超出区域边界时在写入前失败。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldRejectCollisionAndOverflowBeforeWriting()
    {
        var collisionMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" },
                new() { PropertyName = nameof(EntityDynamicLine.Quantity), Title = "Qty" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string", ColumnIndex = 1 }
            }
        };
        var entity = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A", Quantity = 1,
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        var collisionLayout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines,
                region => region.Mapping(collisionMapping)));
        using var collisionOutput = new MemoryStream();
        var collision = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(entity, collisionLayout, collisionOutput));
        Assert.Equal(BingOfficesStage.Plan, collision.Stage);
        Assert.Empty(collisionOutput.ToArray());

        var overflowMapping = new ExcelMappingConfiguration
        {
            Columns = collisionMapping.Columns,
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string" }
            }
        };
        var overflowLayout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines,
                region => region.End("B2").Mapping(overflowMapping)));
        using var overflowOutput = new MemoryStream();
        var overflow = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(entity, overflowLayout, overflowOutput));
        Assert.Equal(BingOfficesStage.Plan, overflow.Stage);
        Assert.Contains("列表区域超出声明边界", overflow.Message);
        Assert.Empty(overflowOutput.ToArray());
    }

    /// <summary>
    /// 验证实体模板导入前检查合并结构并保留模板流。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldValidateMergeAndPreserveTemplateStream()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "B1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Invoice");
            sheet.Cell("A1").Value = "Template";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Range("A1:B1").Merge();
            workbook.SaveAs(template);
        }
        var original = template.ToArray();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice { Title = "Filled" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        Assert.Equal(original, template.ToArray());
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            Assert.Equal("Filled", workbook.Worksheet("Invoice").Cell("A1").GetString());
            Assert.True(workbook.Worksheet("Invoice").Cell("A1").Style.Font.Bold);
            Assert.Single(workbook.Worksheet("Invoice").MergedRanges);
        }

        using var source = new MemoryStream(output.ToArray());
        using var validationTemplate = new MemoryStream(original);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(validationTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Filled", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体异步往返和预取消行为。
    /// </summary>
    [Fact]
    public async Task EntityAsync_ShouldRoundTripAndHonorCancellation()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportEntityAsync(
            new EntityInvoice { Title = "Async" }, layout, output);
        Assert.NotEmpty(output.ToArray());

        using var input = new MemoryStream(output.ToArray());
        var result = await new ClosedXmlExcelImporter().ImportEntityAsync(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Async", result.Entity.Title);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var canceled = new MemoryStream();
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportEntityAsync(
                new EntityInvoice { Title = "Canceled" }, layout, canceled, cancellation.Token));
        Assert.Empty(canceled.ToArray());
    }

    /// <summary>
    /// 验证实体固定单元格使用命名转换器完成双向转换。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseNamedConverter()
    {
        var converter = new EntityTitleConverter();
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title, converter.Name));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter(valueConverters: new[] { converter }).ExportEntity(
            new EntityInvoice { Title = "Sales" }, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("entity:Sales", workbook.Worksheet("Invoice").Cell("A1").GetString());

        using var source = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter(valueConverters: new[] { converter })
            .ImportEntity(source, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Sales", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体固定单元格在双向流程中复用映射值映射。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseMappingValueMapInBothDirections()
    {
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = nameof(EntityInvoice.Title),
            ValueMappings = new List<ExcelValueMappingConfiguration>
            {
                new ExcelValueMappingConfiguration { Text = "Display", Value = "Internal" }
            }
        });
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title, cell => cell.Mapping(mapping)));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice { Title = "Internal" }, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("Display", workbook.Worksheet("Invoice").Cell("A1").GetString());

        using var source = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter().ImportEntity(source, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Internal", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体固定单元格校验错误包含结构化定位信息。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldReturnStructuredValidationError()
    {
        var layout = ExcelEntity.Layout<RequiredEntity>(builder => builder
            .Cell("Invoice", "C4", item => item.Title));
        using var input = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("C4").Value = string.Empty;
            workbook.SaveAs(input);
        }
        input.Position = 0;

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("Invoice", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RequiredEntity.Title), error.PropertyName);
    }

    /// <summary>
    /// 验证动态列可以完成 ClosedXML 导出导入往返。
    /// </summary>
    [Fact]
    public void DynamicColumns_ShouldRoundTripThroughClosedXml()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        };
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<DynamicWorkbook>(workbook =>
            workbook.Sheet("Dynamic", root => root.Rows,
                sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal("Alice", row.Name);
        Assert.Equal("east", row.CustomFields["region"]);
    }

    /// <summary>
    /// 验证关系绑定使用配置的 comparer 连接父子数据。
    /// </summary>
    [Fact]
    public void Relations_ShouldUseConfiguredComparerAndBindChildren()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new RelationParent { Id = "P-1" } })
            .AddSheet("Children", new[] { new RelationChild { ParentId = "p-1", Name = "Line" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<RelationWorkbook>(workbook => workbook
            .Sheet<RelationParent>("Parents", root => root.Parents)
            .Sheet<RelationChild>("Children", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var parent = Assert.Single(result.Workbook.Parents);
        var child = Assert.Single(parent.Children);
        Assert.Equal("Line", child.Name);
    }

    /// <summary>
    /// 验证唯一约束和唯一资源限制错误均被完整报告。
    /// </summary>
    [Fact]
    public void Import_ShouldReportUniqueAndUniqueResourceLimitErrors()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Codes",
            new[] { new UniqueRow { Code = "A" }, new UniqueRow { Code = "A" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .Sheet<UniqueRow>("Codes", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(stream, request);
        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(3, error.RowIndex);

        using var limitedSource = new MemoryStream();
        new ClosedXmlExcelExporter().Export(ExcelExport.Workbook(workbook => workbook.AddSheet("Codes",
            new[] { new UniqueRow { Code = "A" }, new UniqueRow { Code = "B" } })), limitedSource);
        limitedSource.Position = 0;
        var limited = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 })
            .Sheet<UniqueRow>("Codes", root => root.Rows));
        var limitedResult = new ClosedXmlExcelImporter().Import(limitedSource, limited);
        Assert.False(limitedResult.IsSuccess);
        Assert.Contains(limitedResult.Errors, item => item.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证达到最大错误数后导入错误集合被截断。
    /// </summary>
    [Fact]
    public void Import_MaxErrors_ShouldTruncateStructuredErrors()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "error-limit" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 })
            .Sheet<Person>("MissingOne", root => root.People)
            .Sheet<Person>("MissingTwo", root => root.People));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(result.ErrorsTruncated);
        Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, result.Errors[0].Code);
    }

    /// <summary>
    /// 验证属性级合并声明生成真实合并区域。
    /// </summary>
    [Fact]
    public void MergeColumns_ShouldCreateRealMergedRanges()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Merged",
            new[]
            {
                new MergeRow { Group = "A", Name = "One" },
                new MergeRow { Group = "A", Name = "Two" },
                new MergeRow { Group = "B", Name = "Three" }
            }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);

        var ranges = workbook.Worksheet("Merged").MergedRanges.ToArray();
        var range = Assert.Single(ranges);
        Assert.Equal("A2:A3", range.RangeAddress.ToString());
    }

    /// <summary>
    /// 验证 XLS 输入格式在写出前被明确拒绝。
    /// </summary>
    [Fact]
    public void UnsupportedXls_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls)
            .AddSheet("People", new[] { new Person { Name = "invalid" } }));
        using var destination = new MemoryStream();
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, destination));
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证新增 XLSB 枚举在 ClosedXML 中也不会静默降级为 XLSX。
    /// </summary>
    [Fact]
    public void UnsupportedXlsb_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xlsb)
            .AddSheet("People", new[] { new Person { Name = "invalid" } }));
        using var destination = new MemoryStream();
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, destination));
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证资源限制在 ClosedXML 加载前拒绝输入。
    /// </summary>
    [Fact]
    public void ResourceLimit_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "large" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证 ZIP 和 XML 预检限制在 DOM 创建前生效。
    /// </summary>
    [Fact]
    public void XlsxPreflight_ShouldRejectConfiguredZipAndXmlResourceLimits()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "resource" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        var sharedStringsBytes = ReplaceZipEntry(bytes, "xl/sharedStrings.xml",
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + new string('x', 128) + "</sst>");
        var stylesBytes = ReplaceZipEntry(bytes, "xl/styles.xml",
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + new string('x', 128) + "</styleSheet>");
        var scenarios = new[]
        {
            (Name: "entries", Limits: new ExcelResourceLimits { MaxZipEntries = 1 }, Bytes: bytes),
            (Name: "entry-size", Limits: new ExcelResourceLimits { MaxZipEntryUncompressedBytes = 1 }, Bytes: bytes),
            (Name: "total-size", Limits: new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 1 }, Bytes: bytes),
            (Name: "compression-ratio", Limits: new ExcelResourceLimits { MaxZipCompressionRatio = 1 }, Bytes: bytes),
            (Name: "shared-strings", Limits: new ExcelResourceLimits { MaxSharedStringsBytes = 1 }, Bytes: sharedStringsBytes),
            (Name: "styles", Limits: new ExcelResourceLimits { MaxStylesBytes = 1 }, Bytes: stylesBytes),
            (Name: "worksheet", Limits: new ExcelResourceLimits { MaxWorksheetBytes = 1 }, Bytes: bytes),
            (Name: "total-worksheet", Limits: new ExcelResourceLimits { MaxTotalWorksheetBytes = 1 }, Bytes: bytes),
            (Name: "xml-depth", Limits: new ExcelResourceLimits { MaxXmlDepth = 1 }, Bytes: bytes),
            (Name: "xml-characters", Limits: new ExcelResourceLimits { MaxXmlCharacters = 1 }, Bytes: bytes)
        };
        foreach (var scenario in scenarios)
        {
            using var source = new MemoryStream(scenario.Bytes, writable: false);
            var limitedRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
                .ResourceLimits(scenario.Limits)
                .Sheet<Person>("People", root => root.People));
            BingOfficesResourceLimitException exception = null;
            try
            {
                new ClosedXmlExcelImporter().Import(source, limitedRequest);
            }
            catch (BingOfficesResourceLimitException caught)
            {
                exception = caught;
            }

            Assert.NotNull(exception);
            Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
            Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        }
    }

    /// <summary>
    /// 验证 AnnotatedOriginal 失败工作簿保留输入内容并写入错误汇总和批注。
    /// </summary>
    [Fact]
    public void FailureWorkbook_AnnotatedOriginal_ShouldPreserveInputAndWriteDiagnostics()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var data = workbook.Worksheets.Add("Data");
            data.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            data.Cell("A2").Value = "invalid";
            workbook.Worksheets.Add("Preserved").Cell("C3").Value = "keep";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var failureOutput = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureOutput
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        failureOutput.Position = 0;
        using var failureWorkbook = new XLWorkbook(failureOutput);
        Assert.Equal("invalid", failureWorkbook.Worksheet("Data").Cell("A2").GetString());
        Assert.Equal("keep", failureWorkbook.Worksheet("Preserved").Cell("C3").GetString());
        Assert.True(failureWorkbook.Worksheet("Data").Cell("A2").HasComment);
        Assert.Contains(result.Errors.Single().Message,
            failureWorkbook.Worksheet("Data").Cell("A2").GetComment().ToString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Bing.Offices", failureWorkbook.Worksheet("Data").Cell("A2").GetComment().Author);
        Assert.Equal("Code", failureWorkbook.Worksheet("_ImportErrors").Cell("A1").GetString());
        Assert.Equal(2, failureWorkbook.Worksheet("_ImportErrors").Cell("D2").GetValue<int>());
    }

    /// <summary>
    /// 验证失败批注冲突策略符合公共契约。
    /// </summary>
    /// <param name="policy">待验证的冲突或不支持功能处理策略。</param>
    /// <param name="expectedText">预期保留的批注文本或错误文本标记。</param>
    /// <param name="expectedAuthor">预期批注作者。</param>
    /// <param name="shouldThrow">是否预期抛出配置异常。</param>
    [Theory]
    [InlineData(ExcelImportCommentConflictPolicy.Preserve, "existing", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Append, "invalid", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Replace, "invalid", "Bing.Offices", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Fail, null, null, true)]
    public void FailureWorkbook_AnnotatedOriginal_ShouldHonorCommentConflictPolicy(
        ExcelImportCommentConflictPolicy policy, string expectedText, string expectedAuthor, bool shouldThrow)
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            var comment = sheet.Cell("A2").CreateComment();
            comment.Author = "source";
            comment.AddText("existing");
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                CommentConflictPolicy = policy
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        ExcelWorkbookImportResult<ErrorNumberWorkbook> result = null;
        Action action = () => result = new ClosedXmlExcelImporter().Import(source, request);

        if (shouldThrow)
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(action);
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Equal(0, destination.Length);
            return;
        }
        action();
        destination.Position = 0;
        using var output = new XLWorkbook(destination);
        var outputComment = output.Worksheet("Data").Cell("A2").GetComment();
        Assert.Contains(expectedText == "invalid" ? result.Errors.Single().Message : expectedText,
            outputComment.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(expectedAuthor, outputComment.Author);
    }

    /// <summary>
    /// 验证 ErrorRowsOnly 仅复制表头和失败行，并附加来源与错误列。
    /// </summary>
    [Fact]
    public void FailureWorkbook_ErrorRowsOnly_ShouldCopyOnlyFailureRows()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = 42;
            sheet.Cell("A3").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        destination.Position = 0;
        using var output = new XLWorkbook(destination);
        var outputSheet = output.Worksheet("Data");
        Assert.Equal(nameof(ErrorNumberRow.Amount), outputSheet.Cell("A1").GetString());
        Assert.Equal("invalid", outputSheet.Cell("A2").GetString());
        Assert.Equal("__SourceSheet", outputSheet.Cell("B1").GetString());
        Assert.Equal(3, outputSheet.Cell("C2").GetValue<int>());
        Assert.Equal("Data", outputSheet.Cell("B2").GetString());
        Assert.True(outputSheet.Cell("A3").IsEmpty());
        Assert.True(output.TryGetWorksheet("_ImportErrors", out _));
    }

    /// <summary>
    /// 验证异步导入通过真实异步写入提交失败工作簿，且不关闭调用方流。
    /// </summary>
    [Fact]
    public async Task FailureWorkbook_ImportAsync_ShouldUseAsyncDestinationAndKeepStreamsOpen()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        await using var destination = new AsyncOnlyWriteStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = await new ClosedXmlExcelImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(source.CanRead);
        Assert.True(destination.CanWrite);
        using var output = new XLWorkbook(new MemoryStream(destination.ToArray(), writable: false));
        Assert.True(output.TryGetWorksheet("_ImportErrors", out _));
    }

    /// <summary>
    /// 验证失败工作簿序列化超限不会污染调用方目标流。
    /// </summary>
    [Fact]
    public void FailureWorkbook_MaxSerializedBytes_ShouldRejectBeforeDestinationCopy()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                MaxSerializedBytes = 1
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Serialize, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证模板异步导入对不支持部件报告 Import 操作上下文。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateAsync_UnsupportedPart_ShouldReportImportOperation()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.Cell("Invoice", "A1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("A1").Value = "Template";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/chart1.xml");

        using var source = new MemoryStream(template.ToArray(), writable: false);
        template.Position = 0;
        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelImporter().ImportForTemplateAsync(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true)));

        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证异常观察器只接收一次通知且观察器异常不覆盖主异常。
    /// </summary>
    [Fact]
    public void ExceptionObserver_ShouldReceiveClosedXmlFailureOnceAndKeepObserverFailureOutOfBand()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "input" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/failure-workbook-chart.xml");
        source.Position = 0;
        using var failureOutput = new MemoryStream();
        var observer = new RecordingObserver(throwOnObserve: true);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureOutput
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelImporter(exceptionObservers: new[] { observer }).Import(source, request));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.True(exception.Data.Contains(BingOfficesExceptionDispatcher.ObserverFailureKey));
        Assert.Empty(failureOutput.ToArray());
    }

    /// <summary>
    /// 验证异步资源限制失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveAsyncResourceFailureOnce()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "async-resource" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var observer = new RecordingObserver();
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 32 })
            .Sheet<Person>("People", root => root.People));

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter(exceptionObservers: new[] { observer })
                .ImportAsync(new TrackingNonSeekableReadStream(generated.ToArray()), request));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证异步不支持能力失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveAsyncUnsupportedFailureOnce()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(ExcelFormat.Xls)
            .AddSheet("People", new[] { new Person { Name = "unsupported" } }));
        var observer = new RecordingObserver();
        using var destination = new MemoryStream();

        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter(exceptionObservers: new[] { observer })
                .ExportAsync(request, destination));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证文件提交失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveFileCommitFailureOnce()
    {
        var observer = new RecordingObserver();
        var commitException = new BingOfficesFileCommitException("commit failed", provider: "ClosedXML");
        var exporter = new ClosedXmlExcelExporter(
            exceptionObservers: new[] { observer },
            fileExportCommitter: new ThrowingCommitter(commitException));
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "commit" } }));

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            exporter.ExportToFileAsync(request, Path.Combine(Path.GetTempPath(), "closedxml-observer.xlsx")));

        Assert.Equal(1, observer.Count);
        Assert.Same(commitException, exception);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.FileCommit, exception.Operation);
    }

    /// <summary>
    /// 验证异步导入导出成功后保持调用方流打开。
    /// </summary>
    [Fact]
    public async Task StreamOwnership_ShouldKeepCallerStreamsOpenForAsyncImportAndExport()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "owned" } }));
        await using var destination = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportAsync(request, destination);
        Assert.True(destination.CanWrite);

        destination.Position = 0;
        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = await new ClosedXmlExcelImporter().ImportAsync(destination, import);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 验证异步非定位输入可以往返并保持源流打开。
    /// </summary>
    [Fact]
    public async Task ImportAsync_NonSeekableInput_ShouldRoundTripAndKeepSourceOpen()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "non-seekable" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, generated);

        using var source = new TrackingNonSeekableReadStream(generated.ToArray());
        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = await new ClosedXmlExcelImporter().ImportAsync(source, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("non-seekable", Assert.Single(result.Workbook.People).Name);
        Assert.True(source.CanRead);
        Assert.Equal(generated.Length, source.BytesRead);
    }

    /// <summary>
    /// 验证独立工作簿可以并行读写且不共享 DOM 状态。
    /// </summary>
    [Fact]
    public async Task IndependentWorkbooks_ShouldRoundTripInParallelWithoutDomSharing()
    {
        var names = await Task.WhenAll(Enumerable.Range(0, 4).Select(async index =>
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
                new[] { new Person { Name = $"parallel-{index}" } }));
            await using var output = new MemoryStream();
            await new ClosedXmlExcelExporter().ExportAsync(request, output);
            output.Position = 0;
            using var workbook = new XLWorkbook(output);
            return workbook.Worksheet("People").Cell(2, 1).GetString();
        }));

        Assert.Equal(new[] { "parallel-0", "parallel-1", "parallel-2", "parallel-3" }, names);
    }

    /// <summary>
    /// 验证各原生比较运算符的通过与失败边界。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    [Theory]
    [InlineData(XLAllowedValues.WholeNumber)]
    [InlineData(XLAllowedValues.Decimal)]
    [InlineData(XLAllowedValues.TextLength)]
    [InlineData(XLAllowedValues.Date)]
    [InlineData(XLAllowedValues.Time)]
    public void WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics(
        XLAllowedValues allowedValues)
    {
        foreach (var operation in new[]
                 {
                     XLOperator.Between, XLOperator.NotBetween, XLOperator.EqualTo,
                     XLOperator.NotEqualTo, XLOperator.GreaterThan, XLOperator.LessThan,
                     XLOperator.EqualOrGreaterThan, XLOperator.EqualOrLessThan
                 })
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");
            var validation = worksheet.Range("A2").CreateDataValidation();
            ConfigureValidation(validation, allowedValues, operation);
            var (validRaw, validText, invalidRaw, invalidText) = GetComparisonValues(allowedValues, operation);

            var valid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), validRaw,
                validText, CultureInfo.InvariantCulture, isDate1904: allowedValues == XLAllowedValues.Date,
                CancellationToken.None);
            Assert.True(valid.IsValid, $"{allowedValues}/{operation}: {valid.Message}");

            var invalid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), invalidRaw,
                invalidText, CultureInfo.InvariantCulture, isDate1904: allowedValues == XLAllowedValues.Date,
                CancellationToken.None);
            Assert.False(invalid.IsValid, $"{allowedValues}/{operation} must reject its invalid boundary.");
            Assert.False(invalid.IsUnsupported);
        }
    }

    /// <summary>
    /// 验证空值策略和列表区域解析均使用明确的 Workbook 规则边界。
    /// </summary>
    [Fact]
    public void WorkbookValidationBlankAndListRules_ShouldUseExplicitPolicies()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "A";
        lookup.Cell("A2").Value = "B";
        worksheet.Cell("C1").Value = "C";
        worksheet.Cell("C2").Value = "D";

        var blankValidation = worksheet.Range("A2").CreateDataValidation();
        blankValidation.AllowedValues = XLAllowedValues.WholeNumber;
        blankValidation.Operator = XLOperator.EqualTo;
        blankValidation.Value = "1";
        blankValidation.IgnoreBlanks = true;
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), null,
            string.Empty, CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);
        blankValidation.IgnoreBlanks = false;
        var blankRejected = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), null,
            string.Empty, CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.False(blankRejected.IsValid);
        Assert.Equal("不允许 Workbook 校验目标为空。", blankRejected.Message);

        blankValidation.ClearRanges();
        var listValidation = worksheet.Range("A2").CreateDataValidation();
        listValidation.AllowedValues = XLAllowedValues.List;
        listValidation.Value = "Lookup!$A$1:$A$2";
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "B", "B",
            CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);

        listValidation.ClearRanges();
        var localListValidation = worksheet.Range("A2").CreateDataValidation();
        localListValidation.AllowedValues = XLAllowedValues.List;
        localListValidation.Value = "$C$1:$C$2";
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "D", "D",
            CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);
    }

    /// <summary>
    /// 验证无法解析的引用返回不支持结果。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="expression">待验证的原生引用表达式。</param>
    [Theory]
    [InlineData(XLAllowedValues.List, "KnownValues")]
    [InlineData(XLAllowedValues.List, "[External.xlsx]Lookup!A1:A2")]
    [InlineData(XLAllowedValues.List, "MissingSheet!A1:A2")]
    [InlineData(XLAllowedValues.Custom, "MissingName=1")]
    [InlineData(XLAllowedValues.Custom, "MissingSheet!A2=1")]
    public void WorkbookValidationUnresolvableReferences_ShouldBeExplicitlyUnsupported(
        XLAllowedValues allowedValues, string expression)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var validation = worksheet.Range("A2").CreateDataValidation();
        validation.AllowedValues = allowedValues;
        validation.Value = expression;

        var result = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "1", "1",
            CultureInfo.InvariantCulture, false, CancellationToken.None);

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
        Assert.Equal("Workbook Data Validation 规则类型或公式暂不支持。", result.Message);
    }

    /// <summary>
    /// 验证单单元格比较公式保留为受支持的 Custom 规则，并且不依赖当前值回退。
    /// </summary>
    [Fact]
    public void WorkbookValidationSimpleCustomComparison_ShouldValidateResolvedCellValue()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var validation = worksheet.Range("A2").CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.Custom;
        validation.Value = "$A$2>=5";
        worksheet.Cell("A2").Value = "5";

        var valid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "5", "5",
            CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.True(valid.IsValid, valid.Message);

        worksheet.Cell("A2").Value = "4";
        var invalid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "4", "4",
            CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.False(invalid.IsValid);
        Assert.False(invalid.IsUnsupported);
    }

    /// <summary>
    /// 验证模板 leaveOpen 语义及既有样式保留。
    /// </summary>
    [Fact]
    public void TemplateLeaveOpen_ShouldPreserveTemplateStreamAndExistingStyle()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell(1, 1).Value = "Template Header";
            sheet.Cell(1, 1).Style.Font.Italic = true;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Templated" } }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        Assert.True(template.CanRead);
        output.Position = 0;
        using var result = new XLWorkbook(output);
        Assert.Equal("Templated", result.Worksheet("People").Cell(2, 1).GetString());
        Assert.True(result.Worksheet("People").Cell(1, 1).Style.Font.Italic);
    }

    /// <summary>
    /// 验证模板中的公式、批注、验证和条件格式可以保留。
    /// </summary>
    [Fact]
    public void Template_ShouldPreserveFormulaCommentValidationAndConditionalFormatting()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("D1").FormulaA1 = "=1+1";
            sheet.Cell("D1").CreateComment().AddText("template comment");
            var validation = sheet.Range("D2:D10").CreateDataValidation();
            validation.WholeNumber.Between(1, 10);
            var conditional = sheet.Range("E2:E10").AddConditionalFormat();
            conditional.WhenGreaterThan(5).Fill.SetBackgroundColor(XLColor.Yellow);
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Templated" } }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var outputSheet = workbook.Worksheet("People");
        Assert.Equal("1+1", outputSheet.Cell("D1").FormulaA1);
        Assert.True(outputSheet.Cell("D1").HasComment);
        Assert.Equal("template comment", outputSheet.Cell("D1").GetComment().ToString());
        Assert.Single(outputSheet.DataValidations);
        Assert.Single(outputSheet.ConditionalFormats);
    }

    /// <summary>
    /// 验证 leaveOpen=false 时模板流在导出成功后被释放。
    /// </summary>
    [Fact]
    public void TemplateLeaveOpenFalse_ShouldDisposeTemplateAfterExport()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("People").Cell("A1").Value = "Name";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: false)
            .AddSheet("People", new[] { new Person { Name = "disposed" } }));
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, output);

        Assert.False(template.CanRead);
    }

    /// <summary>
    /// 验证预取消的模板操作仍按 leaveOpen=false 释放模板流。
    /// </summary>
    [Fact]
    public async Task TemplateLeaveOpenFalse_ShouldDisposeOnPreCancelledTemplateOperations()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.Cell("Invoice", "A1", item => item.Title));
        using var exportTemplate = new MemoryStream(new byte[] { 1, 2, 3 });
        using var exportDestination = new MemoryStream();
        using var importTemplate = new MemoryStream(new byte[] { 4, 5, 6 });
        using var source = new MemoryStream(new byte[] { 7, 8, 9 });
        using var workbookTemplate = new MemoryStream(new byte[] { 10, 11, 12 });
        using var workbookDestination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(
                new EntityInvoice { Title = "cancelled" }, layout,
                new ExcelEntityTemplateOptions(exportTemplate, leaveOpen: false), exportDestination,
                cancellation.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().ImportForTemplateAsync(
                source, layout, new ExcelEntityTemplateOptions(importTemplate, leaveOpen: false),
                cancellation.Token));
        var workbookRequest = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(workbookTemplate, leaveOpen: false)
            .AddSheet("People", new[] { new Person { Name = "cancelled" } }));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportAsync(workbookRequest, workbookDestination,
                cancellation.Token));

        Assert.False(exportTemplate.CanRead);
        Assert.False(importTemplate.CanRead);
        Assert.False(workbookTemplate.CanRead);
    }

    /// <summary>
    /// 验证包含 Chart 部件的模板在 ClosedXML 加载前失败。
    /// </summary>
    [Fact]
    public void TemplateWithChartPart_ShouldFailBeforeClosedXmlLoad()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            sourceWorkbook.Worksheets.Add("People").Cell(1, 1).Value = "Name";
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/chart1.xml");
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "rejected" } }));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Empty(output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证不支持的模板部件按名称在输出前 fail-fast。
    /// </summary>
    /// <param name="entryName">要注入模板的 ZIP 部件名称。</param>
    [Theory]
    [InlineData("xl/pivotTables/pivotTable1.xml")]
    [InlineData("xl/pivotCache/pivotCacheDefinition1.xml")]
    [InlineData("xl/media/image1.png")]
    [InlineData("xl/tables/table1.xml")]
    [InlineData("xl/vbaProject.bin")]
    [InlineData("xl/vbaProjectSignature.bin")]
    [InlineData("xl/externalLinks/externalLink1.xml")]
    public void UnsupportedTemplateParts_ShouldFailBeforeClosedXmlLoad(string entryName)
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            sourceWorkbook.Worksheets.Add("People").Cell("A1").Value = "Name";
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry(entryName);
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "rejected" } }));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Empty(output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 Provider 使用映射计划工厂并隔离导入导出方向。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_ShouldBeUsedByProviderAndKeepDirectionIsolation()
    {
        var inner = ExcelMappingPlanFactoryProvider.CreateDefault();
        var counting = new CountingMappingPlanFactory(inner);
        var document = new ExcelMappingDocument { UseConventionFallback = true };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "cached", Age = 1, Score = 1m } },
            sheet => sheet.Mapping(document)));
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter(mappingPlanFactory: counting);

        exporter.Export(request, first);
        exporter.Export(request, second);

        Assert.True(counting.CreateCount >= 2);
        var exportPlan = inner.Create<Person>(document, MappingDirection.Export);
        Assert.Same(exportPlan, inner.Create<Person>(document, MappingDirection.Export));
        Assert.NotSame(exportPlan, inner.Create<Person>(document, MappingDirection.Import));
    }

    /// <summary>
    /// 验证映射计划按类型、配置和动态键隔离，并在工厂失败后恢复。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_ShouldIsolateTypeConfigurationAndDynamicKeysAndRecoverAfterFailure()
    {
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault();
        var document = new ExcelMappingDocument { UseConventionFallback = true };
        var basePlan = factory.Create<Person>(document, MappingDirection.Export);
        Assert.Same(basePlan, factory.Create<Person>(document, MappingDirection.Export));
        Assert.NotSame(basePlan, factory.Create<DynamicRow>(document, MappingDirection.Export));

        var configured = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(Person.Name), Title = "Display Name" }
            }
        };
        var configuredPlan = factory.Create<Person>(document, configured, MappingDirection.Export);
        Assert.NotSame(basePlan, configuredPlan);
        Assert.Equal("Display Name", configuredPlan.Columns.Single(column =>
            column.Name == nameof(Person.Name)).Title);

        var dynamicConfiguration = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "region", Title = "Region", DataTypeName = "string" }
            }
        };
        var dynamicPlan = factory.Create<DynamicRow>(document, dynamicConfiguration,
            MappingDirection.Export);
        Assert.Contains(dynamicPlan.DynamicColumns, column => column.Key == "region");
        Assert.NotSame(dynamicPlan, factory.Create<DynamicRow>(document, MappingDirection.Export));

        var invalid = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "", Title = "invalid" }
            }
        };
        Assert.Throws<ArgumentException>(() => factory.Create<Person>(document, invalid,
            MappingDirection.Export));
        var recovered = factory.Create<Person>(document, MappingDirection.Export);
        Assert.Same(basePlan, recovered);
    }

    /// <summary>
    /// 验证模板批注冲突策略产生对应的保留、覆盖或忽略结果。
    /// </summary>
    /// <param name="policy">要验证的批注冲突策略。</param>
    /// <param name="expected">预期读回的批注文本。</param>
    [Theory]
    [InlineData(ExcelCommentConflictPolicy.Preserve, "existing")]
    [InlineData(ExcelCommentConflictPolicy.Replace, "generated")]
    public void TemplateCommentConflict_ShouldHonorPolicy(ExcelCommentConflictPolicy policy, string expected)
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated", Age = 1, Score = 1m } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated", "test"))
                        })
                    })
                    .CommentConflicts(policy)));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var result = new XLWorkbook(output);
        Assert.Equal(expected, result.Worksheet("People").Cell("A1").GetComment().ToString());
    }

    /// <summary>
    /// 验证批注冲突 fail 策略在输出前报告结构化错误。
    /// </summary>
    [Fact]
    public void TemplateCommentConflictFail_ShouldRejectBeforeOutput()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated" } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated"))
                        })
                    })
                    .CommentConflicts(ExcelCommentConflictPolicy.Fail)));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证模板覆盖策略可以清理既有批注和样式。
    /// </summary>
    [Fact]
    public void TemplateOverwrite_ShouldClearTemplateCommentAndStyle()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated" } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .TemplateCellOverwrite(ExcelTemplateCellOverwritePolicy.ReplaceTemplate)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated"))
                        })
                    })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var result = new XLWorkbook(output);
        var cell = result.Worksheet("People").Cell("A1");
        Assert.False(cell.Style.Font.Bold);
        Assert.Equal("generated", cell.GetComment().ToString());
    }

    /// <summary>
    /// 验证重叠的实体列表区域在布局构建阶段被拒绝。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectOverlappingListRegions()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("Lines", "A1", item => item.Lines)
            .ListRegion("Lines", "A2", item => item.Lines)));
    }

    /// <summary>
    /// 验证实体模板缺少工作表或合并区域时被拒绝。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldRejectMissingSheetAndMerge()
    {
        var missingSheetLayout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var missingSheetTemplate = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Other");
            workbook.SaveAs(missingSheetTemplate);
        }
        missingSheetTemplate.Position = 0;
        using var missingSheetOutput = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice(), missingSheetLayout,
                new ExcelEntityTemplateOptions(missingSheetTemplate, leaveOpen: true), missingSheetOutput));

        var missingMergeLayout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "A1", item => item.Title));
        using var missingMergeTemplate = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice");
            workbook.SaveAs(missingMergeTemplate);
        }
        missingMergeTemplate.Position = 0;
        using var missingMergeOutput = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice(), missingMergeLayout,
                new ExcelEntityTemplateOptions(missingMergeTemplate, leaveOpen: true), missingMergeOutput));
    }

    /// <summary>
    /// 验证实体模板异步往返并执行模板结构校验。
    /// </summary>
    [Fact]
    public async Task EntityTemplateAsync_ShouldRoundTripAndValidateTemplate()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("A1").Value = "Template";
            workbook.SaveAs(template);
        }
        var templateBytes = template.ToArray();
        template.Position = 0;
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportForTemplateAsync(
            new EntityInvoice { Title = "Async template" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var source = new MemoryStream(output.ToArray());
        using var validationTemplate = new MemoryStream(templateBytes);
        var result = await new ClosedXmlExcelImporter().ImportForTemplateAsync(source, layout,
            new ExcelEntityTemplateOptions(validationTemplate));

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Async template", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体异步文件导出提交后可以重新读取工作簿。
    /// </summary>
    [Fact]
    public async Task EntityExportToFileAsync_ShouldCommitReadableWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"closedxml-entity-{Guid.NewGuid():N}.xlsx");
        try
        {
            var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
                .Cell("Invoice", "A1", item => item.Title));
            await new ClosedXmlExcelExporter().ExportEntityToFileAsync(
                new EntityInvoice { Title = "file entity" }, layout, path);

            using var workbook = new XLWorkbook(path);
            Assert.Equal("file entity", workbook.Worksheet("Invoice").Cell("A1").GetString());
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证非法 XLSX 输入的预检异常带有 ClosedXML Provider 上下文。
    /// </summary>
    [Fact]
    public void InvalidXlsx_ShouldUseClosedXmlProviderInPreflightError()
    {
        using var source = new MemoryStream(new byte[] { 1, 2, 3, 4 });
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("ClosedXML", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 表示基础导入测试的工作簿模型。
    /// </summary>
    private sealed class PeopleWorkbook
    {
        /// <summary>
        /// 获取人员数据行集合。
        /// </summary>
        public List<Person> People { get; } = new();
    }

    /// <summary>
    /// 表示边界导入测试的工作簿模型。
    /// </summary>
    private sealed class BoundaryWorkbook
    {
        /// <summary>
        /// 获取边界数据行集合。
        /// </summary>
        public List<BoundaryRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示单元格边界导入测试的数据行。
    /// </summary>
    private sealed class BoundaryRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置显式空字符串字段。
        /// </summary>
        public string Empty { get; set; }

        /// <summary>
        /// 获取或设置错误单元格文本字段。
        /// </summary>
        public string Error { get; set; }

        /// <summary>
        /// 获取或设置合并锚点字段。
        /// </summary>
        public string Merged { get; set; }
    }

    /// <summary>
    /// 表示错误数值单元格测试的工作簿模型。
    /// </summary>
    private sealed class ErrorNumberWorkbook
    {
        /// <summary>
        /// 获取错误数值数据行集合。
        /// </summary>
        public List<ErrorNumberRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示错误数值单元格测试的数据行。
    /// </summary>
    private sealed class ErrorNumberRow
    {
        /// <summary>
        /// 获取或设置金额字段。
        /// </summary>
        public int Amount { get; set; }
    }

    /// <summary>
    /// 验证动态列布局、行高和样式优先级在真实 XLSX 中生效。
    /// </summary>
    [Fact]
    public void DynamicLayoutAndRowHeight_ShouldUsePhysicalColumnAndStylePrecedence()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string),
            Placement = ExcelColumnPlacement.Before(nameof(DynamicRow.Name)),
            HeaderStyle = new ExcelCellStyle { Bold = true },
            BodyStyle = new ExcelCellStyle { NumberFormat = "@" }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet
                .DynamicColumns(row => row.CustomFields, new[] { definition })
                .HeaderStyle(new ExcelCellStyle { Italic = true })
                .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00" })
                .RowHeight(new ExcelRowHeightOptions { HeaderHeight = 22, BodyHeight = 18 })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("Dynamic");
        Assert.Equal("Region", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal("east", sheet.Cell("A2").GetString());
        Assert.True(sheet.Cell("A1").Style.Font.Bold);
        Assert.True(sheet.Cell("A1").Style.Font.Italic);
        Assert.Equal("@", sheet.Cell("A2").Style.NumberFormat.Format);
        Assert.Equal(22, sheet.Row(1).Height);
        Assert.Equal(18, sheet.Row(2).Height);
    }

    /// <summary>
    /// 验证动态列物理索引保留最终列位置并拒绝冲突。
    /// </summary>
    [Fact]
    public void DynamicPhysicalIndex_ShouldUseFinalColumnAndRejectConflicts()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            PhysicalColumnIndex = 3
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("Dynamic");
            Assert.Equal("Name", sheet.Cell("A1").GetString());
            Assert.True(sheet.Cell("B1").IsEmpty());
            Assert.True(sheet.Cell("C1").IsEmpty());
            Assert.Equal("Region", sheet.Cell("D1").GetString());
            Assert.Equal("Alice", sheet.Cell("A2").GetString());
            Assert.True(sheet.Cell("B2").IsEmpty());
            Assert.True(sheet.Cell("C2").IsEmpty());
            Assert.Equal("east", sheet.Cell("D2").GetString());
        }

        var conflicting = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice" } }, sheet => sheet.DynamicColumns(row => row.CustomFields,
                new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "first", Title = "First", PhysicalColumnIndex = 2
                    },
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "second", Title = "Second", PhysicalColumnIndex = 2
                    }
                })));
        using var conflictingOutput = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ClosedXmlExcelExporter().Export(conflicting, conflictingOutput));
        Assert.Empty(conflictingOutput.ToArray());
    }

    /// <summary>
    /// 验证映射中的固定列索引保持声明的物理列顺序。
    /// </summary>
    [Fact]
    public void MappingColumnIndex_ShouldPreserveConfiguredFixedColumnOrder()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "Alice", Age = 30 } }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(Person.Age), ColumnIndex = 0 },
                        new() { PropertyName = nameof(Person.Name), ColumnIndex = 1 }
                    }
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Age", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal(30, sheet.Cell("A2").GetDouble());
        Assert.Equal("Alice", sheet.Cell("B2").GetString());
    }

    /// <summary>
    /// 验证固定列、动态列和稀疏列宽使用统一的物理布局。
    /// </summary>
    [Fact]
    public void MixedFixedColumnIndex_ShouldFillRemainingSlotsAndSkipSparseWidthGaps()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new SparseFixedRow { Name = "Alice", Age = 30 } }, sheet => sheet
                .Mapping(new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(SparseFixedRow.Name), ColumnIndex = 1 }
                    }
                })
                .DynamicColumns(row => CreateRegionValues(row), new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "region", Title = "Region", PhysicalColumnIndex = 3
                    }
                })
                .ColumnWidth(new ExcelColumnWidthOptions
                {
                    Mode = ExcelColumnWidthMode.Fixed,
                    FixedWidth = 20
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Age", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal(string.Empty, sheet.Cell("C1").GetString());
        Assert.Equal("Region", sheet.Cell("D1").GetString());
        Assert.Equal(20, sheet.Column("A").Width);
        Assert.Equal(20, sheet.Column("B").Width);
        Assert.NotEqual(20, sheet.Column("C").Width);
        Assert.Equal(20, sheet.Column("D").Width);
    }

    /// <summary>
    /// 创建稀疏固定列测试使用的动态列值。
    /// </summary>
    /// <param name="_">当前数据行；该测试不读取行内容。</param>
    /// <returns>动态列键和值。</returns>
    private static IDictionary<string, object> CreateRegionValues(SparseFixedRow _)
    {
        return new Dictionary<string, object>
        {
            ["region"] = "east"
        };
    }

    /// <summary>
    /// 验证规范化后的重复表头在行物化前被拒绝。
    /// </summary>
    [Fact]
    public void Import_DuplicateNormalizedHeaders_ShouldFailBeforeMaterializingRows()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = " name ";
            sheet.Cell("A2").Value = "first";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
    }

    /// <summary>
    /// 验证工作表、列数和物理单元格资源限制的精确边界。
    /// </summary>
    [Fact]
    public void ResourceLimits_ShouldEnforceSheetColumnAndPhysicalCellCounts()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", new[] { new Person { Name = "A", Age = 1 } })
            .AddSheet("Two", new[] { new Person { Name = "B", Age = 2 } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();

        using var sheetSource = new MemoryStream(bytes, writable: false);
        var sheetRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(sheetSource, sheetRequest));

        using var columnSource = new MemoryStream(bytes, writable: false);
        var columnRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 3 })
            .Sheet<Person>("One", root => root.People));
        var columnResult = new ClosedXmlExcelImporter().Import(columnSource, columnRequest);
        Assert.True(columnResult.IsSuccess, string.Join(";", columnResult.Errors.Select(error => error.Message)));

        using var overColumnSource = new MemoryStream(bytes, writable: false);
        var overColumnRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 2 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(overColumnSource, overColumnRequest));

        var physicalCells = CountPhysicalCells(bytes);
        using var exactCellSource = new MemoryStream(bytes, writable: false);
        var exactCellRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxCells = physicalCells })
            .Sheet<Person>("One", root => root.People));
        var exactCellResult = new ClosedXmlExcelImporter().Import(exactCellSource, exactCellRequest);
        Assert.True(exactCellResult.IsSuccess,
            string.Join(";", exactCellResult.Errors.Select(error => error.Message)));

        using var overCellSource = new MemoryStream(bytes, writable: false);
        var overCellRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxCells = physicalCells - 1 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(overCellSource, overCellRequest));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxSheets = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxCells = -1 }.Validate());
    }

    /// <summary>
    /// 验证同步工作簿准入在队列已满时于 DOM 创建前拒绝请求。
    /// </summary>
    [Fact]
    public void WorkbookAdmission_ShouldRejectWhenQueueCapacityIsExceeded()
    {
        var admission = new ClosedXmlWorkbookAdmission(new ClosedXmlProviderOptions
        {
            MaxConcurrentWorkbooks = 1,
            MaxQueuedOperations = 0
        });
        using var held = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            admission.Acquire(CancellationToken.None, BingOfficesOperation.Import));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
    }

    /// <summary>
    /// 验证异步准入的等待、取消、释放和并发统计行为。
    /// </summary>
    [Fact]
    public async Task WorkbookAdmissionAsync_ShouldWaitCancelAndRelease()
    {
        var admission = new ClosedXmlWorkbookAdmission(new ClosedXmlProviderOptions
        {
            MaxConcurrentWorkbooks = 1,
            MaxQueuedOperations = 1
        });
        using var held = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);
        var waiting = admission.AcquireAsync(CancellationToken.None, BingOfficesOperation.Import).AsTask();
        await Task.Yield();
        Assert.False(waiting.IsCompleted);

        held.Dispose();
        await using (await waiting)
        {
        }

        using var heldAgain = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);
        using var cancellation = new CancellationTokenSource();
        var canceled = admission.AcquireAsync(cancellation.Token, BingOfficesOperation.Import).AsTask();
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => canceled);
        heldAgain.Dispose();

        using var afterCancellation = admission.Acquire(CancellationToken.None, BingOfficesOperation.Import);
        Assert.Equal(1, admission.RequestedConcurrency);
        Assert.Equal(1, admission.MaxQueuedOperations);
        Assert.Equal(1, admission.MaxActualParallelism);
        Assert.True(admission.MaxObservedQueuedOperations >= 1);
    }

    /// <summary>
    /// 验证导出在目标流阻塞时仍能释放 DOM 准入供后续请求使用。
    /// </summary>
    [Fact]
    public async Task ExportAsync_ShouldReleaseAdmissionBeforeBlockedDestinationCopy()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesClosedXml(options =>
            {
                options.MaxConcurrentWorkbooks = 1;
                options.MaxQueuedOperations = 1;
            })
            .BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "blocked" } }));
        await using var blockedDestination = new BlockingWriteStream();

        var first = exporter.ExportAsync(request, blockedDestination);
        await blockedDestination.WriteStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));

        await using var secondDestination = new MemoryStream();
        var second = exporter.ExportAsync(request, secondDestination);
        await second.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(secondDestination.Length > 0);

        blockedDestination.Release();
        await first;
    }

    /// <summary>
    /// 验证多个实体列表区域导入后按关系配置绑定子集合。
    /// </summary>
    [Fact]
    public void EntityRelations_ShouldBindMultipleListRegionsAfterImport()
    {
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var entity = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new EntityRelationParent { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new EntityRelationChild { ParentId = "p-1", Name = "child" }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, output);
        using var input = new MemoryStream(output.ToArray());

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var parent = Assert.Single(result.Entity.Parents);
        Assert.Equal("child", Assert.Single(parent.Children).Name);
    }

    /// <summary>
    /// 验证实体关系绑定不会丢失子项动态列值。
    /// </summary>
    [Fact]
    public void EntityRelationsWithDynamicColumns_ShouldRoundTripDynamicValues()
    {
        var childMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityRelationChild.ParentId), Title = "ParentId" },
                new() { PropertyName = nameof(EntityRelationChild.Name), Title = "Name" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string" }
            }
        };
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children,
                region => region.Mapping(childMapping))
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var entity = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new() { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new()
                {
                    ParentId = "P-1", Name = "child",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var child = Assert.Single(Assert.Single(result.Entity.Parents).Children);
        Assert.Equal("child", child.Name);
        Assert.Equal("North", child.CustomFields["zone"]);
    }

    /// <summary>
    /// 验证关系缺失父项报告结构化错误并遵守异步取消。
    /// </summary>
    [Fact]
    public async Task EntityRelations_ShouldReportMissingParentAndHonorAsyncCancellation()
    {
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var source = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new() { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new() { ParentId = "P-2", Name = "orphan" }
            }
        };
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        var result = new ClosedXmlExcelImporter().ImportEntity(exported, layout);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.Relationship);
        Assert.Empty(Assert.Single(result.Entity.Parents).Children);

        using var canceledSource = new MemoryStream(exported.ToArray());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().ImportEntityAsync(canceledSource, layout, cancellation.Token));
    }

    /// <summary>
    /// 验证提交失败时保留目标目录并清理原子提交临时文件。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldKeepExistingDirectoryAndCleanTemporaryFileOnCommitFailure()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"closedxml-commit-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
                new[] { new Person { Name = "commit-failure" } }));
            var exception = Assert.Throws<BingOfficesFileCommitException>(() =>
                new ClosedXmlExcelExporter().ExportToFile(request, directory));

            Assert.Equal(BingOfficesStage.Commit, exception.Stage);
            Assert.True(Directory.Exists(directory));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(directory),
                Path.GetFileName(directory) + ".*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 表示基础 ClosedXML Provider 测试中的人员数据行。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }

        /// <summary>
        /// 获取或设置分数。
        /// </summary>
        public decimal Score { get; set; }
    }

    /// <summary>
    /// 表示稀疏固定列布局测试的数据行。
    /// </summary>
    private sealed class SparseFixedRow
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 表示公式导出测试的数据行。
    /// </summary>
    private sealed class FormulaRow
    {
        /// <summary>
        /// 获取或设置公式文本。
        /// </summary>
        public string Formula { get; set; }
    }

    /// <summary>
    /// 表示公式导出测试的工作簿模型。
    /// </summary>
    private sealed class FormulaWorkbook
    {
        /// <summary>
        /// 获取公式数据行集合。
        /// </summary>
        public List<FormulaRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示日期系统测试的工作簿模型。
    /// </summary>
    private sealed class DateWorkbook
    {
        /// <summary>
        /// 获取日期数据行集合。
        /// </summary>
        public List<DateRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示日期系统测试的数据行。
    /// </summary>
    private sealed class DateRow
    {
        /// <summary>
        /// 获取或设置日期值。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 表示动态列测试的工作簿模型。
    /// </summary>
    private sealed class DynamicWorkbook
    {
        /// <summary>
        /// 获取动态列数据行集合。
        /// </summary>
        public List<DynamicRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示关系测试的工作簿模型。
    /// </summary>
    private sealed class RelationWorkbook
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示关系测试中的父项。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关系绑定后的子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示关系测试中的子项。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示唯一性测试的工作簿模型。
    /// </summary>
    private sealed class UniqueWorkbook
    {
        /// <summary>
        /// 获取唯一性测试的数据行集合。
        /// </summary>
        public List<UniqueRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示唯一性测试的数据行。
    /// </summary>
    private sealed class UniqueRow
    {
        /// <summary>
        /// 获取或设置唯一编码。
        /// </summary>
        [ExcelUnique]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示动态列测试的数据行。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置动态列值集合。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; } =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 表示合并列测试的数据行。
    /// </summary>
    private sealed class MergeRow
    {
        /// <summary>
        /// 获取或设置分组名称。
        /// </summary>
        [MergeColumns]
        public string Group { get; set; }

        /// <summary>
        /// 获取或设置行名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示实体导出测试的根模型。
    /// </summary>
    private sealed class EntityInvoice
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取或设置客户名称。
        /// </summary>
        public string Customer { get; set; }

        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<EntityLine> Lines { get; set; } = new List<EntityLine>();
    }

    /// <summary>
    /// 表示动态列表区域测试的根模型。
    /// </summary>
    private sealed class EntityDynamicRoot
    {
        /// <summary>
        /// 获取或设置动态明细行集合。
        /// </summary>
        public List<EntityDynamicLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示包含固定列和动态字典列的实体明细行。
    /// </summary>
    private sealed class EntityDynamicLine
    {
        /// <summary>
        /// 获取或设置项目名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置动态列值。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }
    }

    /// <summary>
    /// 表示实体关系测试的根模型。
    /// </summary>
    private sealed class EntityRelationRoot
    {
        /// <summary>
        /// 获取或设置父项列表区域。
        /// </summary>
        public List<EntityRelationParent> Parents { get; set; } = new();

        /// <summary>
        /// 获取或设置子项列表区域。
        /// </summary>
        public List<EntityRelationChild> Children { get; set; } = new();
    }

    /// <summary>
    /// 表示实体关系测试中的父项。
    /// </summary>
    private sealed class EntityRelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关系绑定后的子项集合。
        /// </summary>
        public List<EntityRelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示实体关系测试中的子项。
    /// </summary>
    private sealed class EntityRelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置子项动态列字典。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; } =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 表示实体列表区域中的明细行。
    /// </summary>
    private sealed class EntityLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 表示必填实体字段测试模型。
    /// </summary>
    private sealed class RequiredEntity
    {
        /// <summary>
        /// 获取或设置必填标题。
        /// </summary>
        [ExcelRequired]
        public string Title { get; set; }
    }

    /// <summary>
    /// 提供实体标题的双向命名转换。
    /// </summary>
    private sealed class EntityTitleConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "entity-title";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            var text = context.Value?.ToString() ?? string.Empty;
            value = text.StartsWith("entity:", StringComparison.Ordinal)
                ? text.Substring("entity:".Length) : text;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = "entity:" + (context.Value?.ToString() ?? string.Empty);
            return true;
        }
    }

    /// <summary>
    /// 记录异常通知次数并可选择模拟观察器失败。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 指示观察器是否在收到通知后抛出异常。
        /// </summary>
        private readonly bool _throwOnObserve;

        /// <summary>
        /// 初始化一个 <see cref="RecordingObserver"/> 类型的实例。
        /// </summary>
        /// <param name="throwOnObserve">是否在通知后故意抛出异常。</param>
        public RecordingObserver(bool throwOnObserve = false) => _throwOnObserve = throwOnObserve;

        /// <summary>
        /// 获取已收到的通知次数。
        /// </summary>
        public int Count { get; private set; }

        /// <summary>
        /// 获取最近一次收到的异常。
        /// </summary>
        public BingOfficesException LastException { get; private set; }

        /// <inheritdoc />
        public void Observe(BingOfficesException exception)
        {
            Count++;
            LastException = exception;
            if (_throwOnObserve)
                throw new InvalidOperationException("observer failed");
        }
    }

    /// <summary>
    /// 模拟文件提交失败的提交器。
    /// </summary>
    private sealed class ThrowingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 保存预设的提交异常。
        /// </summary>
        private readonly BingOfficesFileCommitException _exception;

        /// <summary>
        /// 初始化一个 <see cref="ThrowingCommitter"/> 类型的实例。
        /// </summary>
        /// <param name="exception">提交时要抛出的异常。</param>
        public ThrowingCommitter(BingOfficesFileCommitException exception) => _exception = exception;

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
            => throw _exception;

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format) => Task.FromException(_exception);
    }

    /// <summary>
    /// 包装映射计划工厂并记录创建次数。
    /// </summary>
    private sealed class CountingMappingPlanFactory : IExcelMappingPlanFactory
    {
        /// <summary>
        /// 保存实际执行计划创建的内部工厂。
        /// </summary>
        private readonly IExcelMappingPlanFactory _inner;

        /// <summary>
        /// 初始化一个 <see cref="CountingMappingPlanFactory"/> 类型的实例。
        /// </summary>
        /// <param name="inner">实际映射计划工厂。</param>
        public CountingMappingPlanFactory(IExcelMappingPlanFactory inner) =>
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));

        /// <summary>
        /// 获取已调用的计划创建次数。
        /// </summary>
        public int CreateCount { get; private set; }

        /// <inheritdoc />
        public IExcelMappingPlan Create<T>(ExcelMappingDocument document, MappingDirection direction)
            where T : class, new()
        {
            CreateCount++;
            return _inner.Create<T>(document, direction);
        }

        /// <inheritdoc />
        public IExcelMappingPlan Create<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
            where T : class, new()
        {
            CreateCount++;
            return _inner.Create<T>(document, requestConfiguration, direction);
        }

        /// <inheritdoc />
        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new() =>
            _inner.CreateWorkbook<T>(document, direction, sheetNames);

        /// <inheritdoc />
        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction,
            IReadOnlyList<string> sheetNames) where T : class, new() =>
            _inner.CreateWorkbook<T>(document, requestConfiguration, direction, sheetNames);
    }

    /// <summary>
    /// 提供不可定位读取流并记录实际读取字节数。
    /// </summary>
    private sealed class TrackingNonSeekableReadStream : Stream
    {
        /// <summary>
        /// 保存底层只读内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="TrackingNonSeekableReadStream"/> 类型的实例。
        /// </summary>
        /// <param name="bytes">要读取的输入字节。</param>
        public TrackingNonSeekableReadStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

        /// <summary>
        /// 获取已读取的字节数。
        /// </summary>
        public long BytesRead { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => true;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            var read = _inner.Read(buffer.Span);
            BytesRead += read;
            return ValueTask.FromResult(read);
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => Task.FromResult(Read(buffer, offset, count));

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    /// <summary>
    /// 只允许异步写入的内存流，用于拒绝异步 API 内的同步写入回退。
    /// </summary>
    private sealed class AsyncOnlyWriteStream : Stream
    {
        /// <summary>
        /// 保存异步写入内容的内存流，由测试包装流负责释放。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 获取已写入内容的字节副本。
        /// </summary>
        /// <returns>当前写入内容的独立字节数组。</returns>
        internal byte[] ToArray() => _inner.ToArray();

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => throw new NotSupportedException();

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
            => _inner.WriteAsync(buffer, cancellationToken);

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
            => _inner.WriteAsync(buffer, offset, count, cancellationToken);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
            => throw new NotSupportedException("不允许同步写入。");

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value)
            => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供可控阻塞的写入流以测试 DOM 准入释放时机。
    /// </summary>
    private sealed class BlockingWriteStream : Stream
    {
        /// <summary>
        /// 保存底层内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 获取首次写入已开始的通知。
        /// </summary>
        public TaskCompletionSource<object> WriteStarted { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 获取允许写入继续执行的内部信号。
        /// </summary>
        private TaskCompletionSource<object> ReleaseSignal { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 释放阻塞的写入操作。
        /// </summary>
        public void Release() => ReleaseSignal.TrySetResult(null);

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);

        /// <inheritdoc />
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            WriteStarted.TrySetResult(null);
            await ReleaseSignal.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
            => WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
            => _inner.Write(buffer, offset, count);

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 将工作簿日期系统标记为 1904 系统并保持输入流可继续使用。
    /// </summary>
    /// <param name="stream">待修改的 XLSX 内存流。</param>
    private static void MarkWorkbookDate1904(MemoryStream stream)
    {
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            Assert.NotNull(entry);
            XDocument document;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true))
                document = XDocument.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var root = document.Root;
            Assert.NotNull(root);
            var workbookProperties = root.Element(root.Name.Namespace + "workbookPr");
            if (workbookProperties == null)
            {
                workbookProperties = new XElement(root.Name.Namespace + "workbookPr");
                root.AddFirst(workbookProperties);
            }
            workbookProperties.SetAttributeValue("date1904", "1");
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("xl/workbook.xml").Open(),
                new UTF8Encoding(false));
            document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        stream.Position = 0;
    }

    /// <summary>
    /// 为管线直接测试配置与 ClosedXML 保存格式一致的原生比较规则。
    /// </summary>
    /// <param name="validation">待配置的原生校验规则。</param>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="operation">待验证的比较运算符。</param>
    private static void ConfigureValidation(IXLDataValidation validation, XLAllowedValues allowedValues,
        XLOperator operation)
    {
        validation.AllowedValues = allowedValues;
        validation.Operator = operation;
        validation.MinValue = allowedValues == XLAllowedValues.Time ? "0.5" : "5";
        validation.MaxValue = allowedValues == XLAllowedValues.Time ? "0.75" : "10";
        validation.Value = validation.MinValue;
        validation.IgnoreBlanks = false;
    }

    /// <summary>
    /// 创建指定比较规则的通过值和失败值。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="operation">待验证的比较运算符。</param>
    /// <returns>分别用于通过与失败场景的原始值及文本表示。</returns>
    private static (object ValidRaw, string ValidText, object InvalidRaw, string InvalidText) GetComparisonValues(
        XLAllowedValues allowedValues, XLOperator operation)
    {
        var (valid, invalid) = operation switch
        {
            XLOperator.Between => (5m, 11m),
            XLOperator.NotBetween => (11m, 5m),
            XLOperator.EqualTo => (5m, 6m),
            XLOperator.NotEqualTo => (6m, 5m),
            XLOperator.GreaterThan => (6m, 5m),
            XLOperator.LessThan => (4m, 5m),
            XLOperator.EqualOrGreaterThan => (5m, 4m),
            XLOperator.EqualOrLessThan => (5m, 6m),
            _ => throw new ArgumentOutOfRangeException(nameof(operation), operation, null)
        };
        if (allowedValues == XLAllowedValues.Time)
        {
            var timeValid = DateTime.MinValue.Add(TimeSpan.FromDays((double)(valid / 10m)));
            var timeInvalid = DateTime.MinValue.Add(TimeSpan.FromDays((double)(invalid / 10m)));
            return (timeValid, timeValid.ToString("HH:mm:ss", CultureInfo.InvariantCulture), timeInvalid,
                timeInvalid.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
        }
        if (allowedValues == XLAllowedValues.Date)
        {
            var dateValid = new DateTime(1904, 1, 1).AddDays((double)valid);
            var dateInvalid = new DateTime(1904, 1, 1).AddDays((double)invalid);
            return (dateValid, dateValid.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), dateInvalid,
                dateInvalid.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        }
        if (allowedValues == XLAllowedValues.TextLength)
        {
            var validText = new string('a', (int)valid);
            var invalidText = new string('a', (int)invalid);
            return (validText, validText, invalidText, invalidText);
        }
        return (valid, valid.ToString(CultureInfo.InvariantCulture), invalid,
            invalid.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// 替换 XLSX ZIP 中指定名称的 XML 部件。
    /// </summary>
    /// <param name="source">原始 XLSX 字节。</param>
    /// <param name="entryName">要替换的 ZIP 部件名称。</param>
    /// <param name="content">新的 UTF-8 部件内容。</param>
    /// <returns>替换部件后的 XLSX 字节。</returns>
    private static byte[] ReplaceZipEntry(byte[] source, string entryName, string content)
    {
        using var stream = new MemoryStream();
        stream.Write(source, 0, source.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry(entryName)?.Delete();
            var entry = archive.CreateEntry(entryName);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
        return stream.ToArray();
    }

    /// <summary>
    /// 流式统计 XLSX 工作表 XML 中的物理单元格元素数量。
    /// </summary>
    /// <param name="source">待统计的 XLSX 字节。</param>
    /// <returns>所有工作表中物理单元格元素的总数。</returns>
    private static long CountPhysicalCells(byte[] source)
    {
        using var stream = new MemoryStream(source, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        long count = 0;
        foreach (var entry in archive.Entries.Where(item =>
                     item.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                     && item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            using var reader = System.Xml.XmlReader.Create(entry.Open(), new System.Xml.XmlReaderSettings
            {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element
                    && string.Equals(reader.LocalName, "c", StringComparison.Ordinal))
                    count++;
            }
        }
        return count;
    }
}
