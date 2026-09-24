using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Internals;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.MiniExcel.Tests;

/// <summary>
/// MiniExcel Provider 的真实 XLSX 职责测试。
/// </summary>
public sealed class MiniExcelProviderTest
{
    /// <summary>
    /// 验证 MiniExcel 对实体布局入口应在读取输入前报告不支持能力。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldFailFastBeforeIo()
    {
        var layout = ExcelEntity.Layout<Person>(builder => builder.Cell("People", "A1", item => item.Name));
        using var source = new MemoryStream(new byte[] { 1, 2, 3 });

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelImporter().ImportEntity(source, layout));

        Assert.Equal("MiniExcel", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证 MiniExcel 实体导出在写出任何字节前报告不支持能力。
    /// </summary>
    [Fact]
    public void EntityExport_ShouldFailFastBeforeOutput()
    {
        var layout = ExcelEntity.Layout<Person>(builder => builder.Cell("People", "A1", item => item.Name));
        using var destination = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().ExportEntity(new Person { Name = "unsupported" }, layout, destination));

        Assert.Equal("MiniExcel", exception.Provider);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证往返应支持多个工作表并映射。
    /// </summary>
    [Fact]
    public void RoundTrip_ShouldSupportMultipleSheetsAndMapping()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30 } })
            .AddSheet("Teams", new[] { new Team { Name = "Core" } }));
        using var stream = new MemoryStream();

        new MiniExcelExcelExporter().Export(request, stream);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<Workbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People)
            .Sheet<Team>("Teams", root => root.Teams));
        var result = new MiniExcelExcelImporter().Import(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.People);
        Assert.Equal("Alice", result.Workbook.People[0].Name);
        Assert.Equal(30, result.Workbook.People[0].Age);
        Assert.Equal("Core", Assert.Single(result.Workbook.Teams).Name);
    }

    /// <summary>
    /// 验证异步往返会使用 MiniExcel 异步 API。
    /// </summary>
    [Fact]
    public async Task AsyncRoundTrip_ShouldUseMiniExcelAsyncApi()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Async", Age = 7 } }));
        using var stream = new MemoryStream();
        await new MiniExcelExcelExporter().ExportAsync(request, stream);
        Assert.True(stream.CanWrite);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<Workbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = await new MiniExcelExcelImporter().ImportAsync(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Async", Assert.Single(result.Workbook.People).Name);
        Assert.True(stream.CanRead);
    }

    /// <summary>
    /// 验证添加 BingOfficesMiniExcel 会注册 Provider 服务。
    /// </summary>
    [Fact]
    public void AddBingOfficesMiniExcel_ShouldRegisterProviderServices()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();

        Assert.IsType<MiniExcelExcelExporter>(provider.GetRequiredService<IExcelExporter>());
        Assert.IsType<MiniExcelExcelImporter>(provider.GetRequiredService<IExcelImporter>());
    }

    /// <summary>
    /// 验证XLS应被拒绝为不支持。
    /// </summary>
    [Fact]
    public void Xls_ShouldBeRejectedAsUnsupported()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(ExcelFormat.Xls)
            .AddSheet("People", new[] { new Person { Name = "XLS" } }));

        Assert.Throws<Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException>(() =>
        {
            using var stream = new MemoryStream();
            new MiniExcelExcelExporter().Export(request, stream);
        });
    }

    /// <summary>
    /// 验证不支持呈现选项应失败之前写入。
    /// </summary>
    [Fact]
    public void UnsupportedPresentationOptions_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "styled", Age = 1 } },
                sheet => sheet.SheetStyle(new Bing.Offices.Styles.ExcelCellStyle { Bold = true })));
        using var stream = new MemoryStream();

        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().Export(request, stream));
        Assert.Equal(0, stream.Length);
    }

    /// <summary>
    /// 验证动态列应往返使用命名转换器。
    /// </summary>
    [Fact]
    public void DynamicColumns_ShouldRoundTripWithNamedConverter()
    {
        var converter = new UpperRegionConverter();
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string),
            ConverterName = "upper-region"
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new DynamicRow { Name = "A", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } }, new DynamicRow { Name = "B", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "west"
            } } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter(new IExcelValueConverter[] { converter })
            .Export(exportRequest, stream);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<DynamicWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet =>
                sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        var result = new MiniExcelExcelImporter(valueConverters:
                new IExcelValueConverter[] { new UpperRegionConverter() })
            .Import(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "EAST", "WEST" }, result.Workbook.Rows
            .Select(row => (string)row.CustomFields["region"]).ToArray());
        Assert.Equal(new[] { 2, 3 }, converter.ExportRows.Distinct().OrderBy(index => index).ToArray());
    }

    /// <summary>
    /// 验证固定列转换器应接收一个基准数据行。
    /// </summary>
    [Fact]
    public void FixedColumnConverter_ShouldReceiveOneBasedDataRows()
    {
        var converter = new RowIndexConverter();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new Person { Name = "first", Age = 1 }, new Person { Name = "second", Age = 2 } },
            sheet => sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns =
                {
                    new ExcelColumnConfiguration
                    {
                        PropertyName = nameof(Person.Name),
                        ConverterName = "row-index"
                    }
                }
            })));
        using var stream = new MemoryStream();

        new MiniExcelExcelExporter(new IExcelValueConverter[] { converter }).Export(request, stream);

        Assert.Equal(new[] { 2, 3 }, converter.ExportRows.Distinct().OrderBy(index => index).ToArray());
    }

    /// <summary>
    /// 验证固定列在映射顺序不一致时仍使用物理列索引。
    /// </summary>
    [Fact]
    public void FixedColumns_ShouldUsePhysicalColumnIndexWhenMappingOrderDiffers()
    {
        var exportMapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(FixedColumnContextSourceRow.Name),
                    Title = "Name",
                    ColumnIndex = 0
                },
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(FixedColumnContextSourceRow.Age),
                    Title = "Age",
                    ColumnIndex = 1
                }
            }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new FixedColumnContextSourceRow { Name = "valid", Age = "not-an-int" } },
            sheet => sheet.Mapping(exportMapping)));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);

        var importMapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(FixedColumnContextRow.Age),
                    Title = "Age",
                    ColumnIndex = 0
                },
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(FixedColumnContextRow.Name),
                    Title = "Name",
                    ColumnIndex = 1,
                    ConverterName = "capture-fixed-context",
                    ValidationRuleNames = { "capture-fixed-context-validation" }
                }
            }
        };
        var converter = new FixedColumnContextConverter();
        var validation = new FixedColumnContextValidationRule();
        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<FixedColumnContextWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue).Mapping(importMapping)));
        var result = new MiniExcelExcelImporter(
                valueConverters: new IExcelValueConverter[] { converter },
                namedValidationRules: new INamedExcelValidationRule[] { validation })
            .Import(stream, importRequest);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var converterContext = Assert.Single(converter.ImportContexts);
        Assert.Equal(2, converterContext.RowIndex);
        Assert.Equal(1, converterContext.ColumnIndex);
        var validationContext = Assert.Single(validation.Contexts);
        Assert.Equal(2, validationContext.RowIndex);
        Assert.Equal(1, validationContext.ColumnIndex);

        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        Assert.Equal(nameof(FixedColumnContextRow.Age), error.PropertyName);
    }

    /// <summary>
    /// 验证动态列应使用物理列索引用于转换器验证并错误。
    /// </summary>
    [Fact]
    public void DynamicColumns_ShouldUsePhysicalColumnIndexForConverterValidationAndErrors()
    {
        var exportDefinitions = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "region",
                Title = "Region",
                DataType = typeof(string)
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "amount",
                Title = "Amount",
                DataType = typeof(string)
            }
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[]
            {
                new DynamicContextRow
                {
                    Code = 42,
                    CustomFields = new Dictionary<string, object>
                    {
                        ["region"] = "invalid",
                        ["amount"] = "not-an-int"
                    }
                }
            }, sheet => sheet.DynamicColumns(row => row.CustomFields, exportDefinitions)));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);

        var importDefinitions = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "region",
                Title = "Region",
                DataType = typeof(string),
                ValidatorName = "reject-invalid-dynamic",
                ConverterName = "capture-dynamic-context"
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "amount",
                Title = "Amount",
                DataType = typeof(int)
            }
        };
        var converter = new DynamicColumnContextConverter();
        var validation = new DynamicColumnContextValidationRule();
        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<DynamicContextWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows,
                sheet => sheet.DynamicColumns(row => row.CustomFields, importDefinitions)));
        var result = new MiniExcelExcelImporter(
                valueConverters: new IExcelValueConverter[] { converter },
                namedValidationRules: new INamedExcelValidationRule[] { validation })
            .Import(stream, importRequest);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var converterContext = Assert.Single(converter.ImportContexts.Where(context => context.RowIndex == 2));
        Assert.Equal(2, converterContext.RowIndex);
        Assert.Equal(2, converterContext.ColumnIndex);
        Assert.All(converter.ImportContexts, context => Assert.Equal(2, context.ColumnIndex));
        var validationContext = Assert.Single(validation.Contexts.Where(context => context.RowIndex == 2));
        Assert.Equal(2, validationContext.RowIndex);
        Assert.Equal(2, validationContext.ColumnIndex);

        var rowErrors = result.Errors.Where(error => error.RowIndex == 2).ToArray();
        Assert.Equal(2, rowErrors.Length);
        var validationError = Assert.Single(rowErrors.Where(error =>
            error.Code == ExcelImportErrorCode.Validation));
        Assert.Equal(2, validationError.RowIndex);
        Assert.Equal(2, validationError.ColumnIndex);
        Assert.Equal("region", validationError.PropertyName);
        var conversionError = Assert.Single(rowErrors.Where(error =>
            error.Code == ExcelImportErrorCode.ValueConversion));
        Assert.Equal(2, conversionError.RowIndex);
        Assert.Equal(3, conversionError.ColumnIndex);
        Assert.Equal("amount", conversionError.PropertyName);
    }

    /// <summary>
    /// 验证验证模式应禁用已配置规则并拒绝工作簿规则。
    /// </summary>
    [Fact]
    public void ValidationMode_ShouldDisableConfiguredRulesAndRejectWorkbookRules()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new ValidationModeRow { Code = string.Empty } }));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, source);

        source.Position = 0;
        var configuredRequest = ExcelImport.Workbook<ValidationModeWorkbook>(workbook =>
            workbook.Sheet<ValidationModeRow>("Data", root => root.Rows));
        var configured = new MiniExcelExcelImporter().Import(source, configuredRequest);
        Assert.False(configured.IsSuccess);
        Assert.Empty(configured.Workbook.Rows);
        Assert.Contains(configured.Errors, error => error.Code == ExcelImportErrorCode.Validation);

        source.Position = 0;
        var disabledRequest = ExcelImport.Workbook<ValidationModeWorkbook>(workbook =>
            workbook.ValidationMode(ExcelImportValidationMode.Disabled)
                .Sheet<ValidationModeRow>("Data", root => root.Rows));
        var disabled = new MiniExcelExcelImporter().Import(source, disabledRequest);
        Assert.True(disabled.IsSuccess, string.Join(";", disabled.Errors.Select(error => error.Message)));
        Assert.Single(disabled.Workbook.Rows);
        Assert.Null(disabled.Workbook.Rows[0].Code);

        foreach (var mode in new[]
        {
            ExcelImportValidationMode.WorkbookRules,
            ExcelImportValidationMode.ConfiguredAndWorkbook
        })
        {
            source.Position = 0;
            var unsupportedRequest = ExcelImport.Workbook<ValidationModeWorkbook>(workbook =>
                workbook.ValidationMode(mode).Sheet<ValidationModeRow>("Data", root => root.Rows));
            Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                new MiniExcelExcelImporter().Import(source, unsupportedRequest));
        }
    }

    /// <summary>
    /// 验证值映射应转换已配置文本到目标值。
    /// </summary>
    [Fact]
    public void ValueMap_ShouldConvertConfiguredTextToTargetValue()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new ValueMapRow { Status = 1 } }));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, source);

        source.Position = 0;
        var request = ExcelImport.Workbook<ValueMapWorkbook>(workbook =>
            workbook.Sheet<ValueMapRow>("Data", root => root.Rows));
        var result = new MiniExcelExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(1, Assert.Single(result.Workbook.Rows).Status);
    }

    /// <summary>
    /// 验证唯一性应拒绝重复行并保留首个行。
    /// </summary>
    [Fact]
    public void Unique_ShouldRejectDuplicateRowsAndPreserveFirstRow()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", new[]
        {
            new UniqueRow { Code = "A-1" },
            new UniqueRow { Code = "A-1" }
        }));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, source);

        source.Position = 0;
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook =>
            workbook.Sheet<UniqueRow>("Data", root => root.Rows));
        var result = new MiniExcelExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Workbook.Rows);
        var error = Assert.Single(result.Errors.Where(item => item.Code == ExcelImportErrorCode.Validation));
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(2, error.FirstRowNumber);
    }

    /// <summary>
    /// 验证日期序列应使用核心日期契约并标志。
    /// </summary>
    [Fact]
    public void DateSerial_ShouldUseCoreDateContractAnd1904Flag()
    {
        var date = MiniExcelValueAdapter.ConvertRaw(0d, "0", typeof(DateTime),
            CultureInfo.InvariantCulture, isDate1904: true);
        Assert.Throws<InvalidCastException>(() => MiniExcelValueAdapter.ConvertRaw(0d, "0",
            typeof(DateTimeOffset), CultureInfo.InvariantCulture, isDate1904: true));
        var offsetAttribute = new ExcelDateAttribute
        {
            OffsetPolicy = ExcelDateOffsetPolicy.UseFixedOffset,
            OffsetMinutes = 480
        };
        var offset = MiniExcelValueAdapter.ConvertRaw(0d, "0", typeof(DateTimeOffset),
            CultureInfo.InvariantCulture, isDate1904: true, dateAttribute: offsetAttribute);
        Assert.IsType<DateTime>(date);
        Assert.Equal(new DateTime(1904, 1, 1), date);
        Assert.IsType<DateTimeOffset>(offset);
        Assert.Equal(TimeSpan.FromHours(8), ((DateTimeOffset)offset).Offset);
        Assert.Equal(new DateTimeOffset(1904, 1, 1, 0, 0, 0, TimeSpan.FromHours(8)), offset);
    }

    /// <summary>
    /// 验证日期值应规范化跨目标类型通过核心契约。
    /// </summary>
    [Fact]
    public void DateValues_ShouldNormalizeCrossTargetTypesThroughCoreContract()
    {
        var localDate = DateTime.SpecifyKind(new DateTime(2026, 9, 16, 12, 30, 0), DateTimeKind.Local);
        var offsetDate = new DateTimeOffset(2026, 9, 16, 12, 30, 0, TimeSpan.FromHours(8));

        Assert.Throws<InvalidCastException>(() => MiniExcelValueAdapter.ConvertRaw(localDate,
            localDate.ToString("O"), typeof(DateTimeOffset), CultureInfo.InvariantCulture));
        var dateTimeOffset = MiniExcelValueAdapter.ConvertRaw(localDate, localDate.ToString("O"),
            typeof(DateTimeOffset), CultureInfo.InvariantCulture, dateAttribute: new ExcelDateAttribute
            {
                OffsetPolicy = ExcelDateOffsetPolicy.UseFixedOffset,
                OffsetMinutes = 480
            });
        var dateTime = MiniExcelValueAdapter.ConvertRaw(offsetDate, offsetDate.ToString("O"),
            typeof(DateTime), CultureInfo.InvariantCulture);

        Assert.Equal(new DateTimeOffset(2026, 9, 16, 12, 30, 0, TimeSpan.FromHours(8)), dateTimeOffset);
        Assert.Equal(new DateTime(2026, 9, 16, 12, 30, 0, DateTimeKind.Unspecified), dateTime);
    }

    /// <summary>
    /// 验证XLSX日期预检应读取工作簿属性并恢复位置。
    /// </summary>
    [Fact]
    public void XlsxDate1904Preflight_ShouldReadWorkbookPropertyAndRestorePosition()
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        using (var writer = new StreamWriter(archive.CreateEntry("xl/workbook.xml").Open(),
                   new UTF8Encoding(false)))
        {
            writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><workbookPr date1904=\"1\"/></workbook>");
        }

        stream.Position = Math.Min(3, stream.Length);
        var originalPosition = stream.Position;

        Assert.True(MiniExcelXlsxPreflight.GetDate1904(stream));
        Assert.Equal(originalPosition, stream.Position);
    }

    /// <summary>
    /// 验证真实 XLSX 夹具会通过 MiniExcel 导入预检。
    /// </summary>
    [Fact]
    public void RealXlsx1904Fixture_ShouldPassMiniExcelImportPreflight()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new Person { Name = "date-system", Age = 1 } }));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, source);
        MarkWorkbookAsDate1904(source);

        source.Position = 0;
        var request = ExcelImport.Workbook<Workbook>(workbook =>
            workbook.Sheet<Person>("Data", root => root.People));
        var result = new MiniExcelExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("date-system", Assert.Single(result.Workbook.People).Name);
    }

    /// <summary>
    /// 验证真实XLSX日期系统应使用核心序列契约。
    /// </summary>
    [Fact]
    public void RealXlsxDateSystems_ShouldUseCoreSerialContract()
    {
        var expected1900 = new[]
        {
            new DateTime(1899, 12, 28, 18, 0, 0), new DateTime(1899, 12, 29, 18, 0, 0),
            new DateTime(1899, 12, 30), new DateTime(1899, 12, 30, 18, 0, 0),
            new DateTime(1899, 12, 31), new DateTime(1900, 1, 1),
            new DateTime(1900, 2, 28), new DateTime(1900, 2, 28),
            new DateTime(1900, 3, 1)
        };
        var expected1904 = new[]
        {
            new DateTime(1903, 12, 29, 18, 0, 0), new DateTime(1903, 12, 30, 18, 0, 0),
            new DateTime(1903, 12, 31), new DateTime(1903, 12, 31, 18, 0, 0),
            new DateTime(1904, 1, 1), new DateTime(1904, 1, 2),
            new DateTime(1904, 2, 29), new DateTime(1904, 3, 1),
            new DateTime(1904, 3, 2)
        };

        foreach (var date1904 in new[] { false, true })
        {
            using var source = CreateDateSerialFixture(date1904);
            var fixedOffsetConverter = new FixedOffsetDateConverter();
            var mapping = new ExcelMappingConfiguration
            {
                Columns =
                {
                    new ExcelColumnConfiguration
                    {
                        PropertyName = nameof(DateRow.Offset),
                        ConverterName = fixedOffsetConverter.Name
                    }
                }
            };
            var request = ExcelImport.Workbook<DateWorkbook>(workbook =>
                workbook.Sheet<DateRow>("Data", root => root.Rows,
                    sheet => sheet.Mapping(mapping)));
            var expectedDates = date1904 ? expected1904 : expected1900;
            var expectedOffsets = expectedDates
                .Select(date => new DateTimeOffset(date, TimeSpan.FromHours(8)))
                .ToArray();

            var result = new MiniExcelExcelImporter(
                valueConverters: new IExcelValueConverter[] { fixedOffsetConverter })
                .Import(source, request);

            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(expectedDates, result.Workbook.Rows.Select(row => row.Date).ToArray());
            Assert.Equal(expectedOffsets, result.Workbook.Rows.Select(row => row.Offset).ToArray());

        }
    }

    /// <summary>
    /// 测试 - 跳过正文行时，MiniExcel Provider 应继续使用实际物理行的日期 serial、上下文和来源行号。
    /// </summary>
    /// <param name="date1904">是否启用 1904 日期系统；false 表示 1900 日期系统。</param>
    /// <param name="headerRowIndex">表头的零基物理行索引。</param>
    /// <param name="skippedDataRows">表头后跳过的数据行数。</param>
    [Theory]
    [InlineData(false, 0, 1)]
    [InlineData(true, 0, 1)]
    [InlineData(false, 2, 2)]
    [InlineData(true, 2, 2)]
    public void DataRowStartIndex_ShouldAlignRawDateSerials(bool date1904,
        int headerRowIndex, int skippedDataRows)
    {
        var serials = new[] { 61d, -0.25d, -1.25d, 1d };
        var dataRowStartIndex = headerRowIndex + 1 + skippedDataRows;
        using var source = CreateSkippedDateSerialFixture(date1904, headerRowIndex, serials);
        var fixedOffsetConverter = new FixedOffsetDateConverter();
        var fixedValidation = new FixedColumnContextValidationRule();
        var dynamicValidation = new DynamicColumnContextValidationRule();
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(SkippedDateRow.Offset),
                    ConverterName = fixedOffsetConverter.Name,
                    ValidationRuleNames = { fixedValidation.Name }
                }
            }
        };
        var dynamicDefinition = new ExcelDynamicColumnDefinition
        {
            Key = "dynamicDate",
            Title = "DynamicDate",
            DataType = typeof(DateTime),
            ValidatorName = dynamicValidation.Name
        };
        var request = ExcelImport.Workbook<SkippedDateWorkbook>(workbook =>
            workbook.Sheet<SkippedDateRow>("Data", root => root.Rows, sheet => sheet
                .HeaderRowIndex(headerRowIndex)
                .DataRowStartIndex(dataRowStartIndex)
                .Mapping(mapping)
                .DynamicColumns(row => row.CustomFields, new[] { dynamicDefinition })));
        var expectedDates = GetSkippedDateExpectations(date1904, serials, skippedDataRows);
        var expectedOffsets = expectedDates
            .Select(date => new DateTimeOffset(date, TimeSpan.FromHours(8)))
            .ToArray();
        var expectedRows = Enumerable.Range(dataRowStartIndex, serials.Length - skippedDataRows).ToArray();
        var expectedFixedContexts = expectedRows
            .Select(row => (RowIndex: row + 1, ColumnIndex: 2))
            .ToArray();
        var expectedDynamicContexts = expectedRows
            .Select(row => (RowIndex: row + 1, ColumnIndex: 3))
            .ToArray();

        var importer = new MiniExcelExcelImporter(valueConverters: new IExcelValueConverter[] { fixedOffsetConverter },
            namedValidationRules: new INamedExcelValidationRule[] { fixedValidation, dynamicValidation });
        fixedOffsetConverter.ImportContexts.Clear();
        fixedValidation.Contexts.Clear();
        dynamicValidation.Contexts.Clear();
        var result = importer.Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Empty(result.Errors);
        Assert.Equal(expectedDates, result.Workbook.Rows.Select(row => row.Date).ToArray());
        Assert.Equal(expectedOffsets, result.Workbook.Rows.Select(row => row.Offset).ToArray());
        Assert.Equal(expectedDates, result.Workbook.Rows
            .Select(row => (DateTime)row.CustomFields["dynamicDate"]).ToArray());
        Assert.Equal(expectedRows, result.Sheets.Single().SourceRows);
        Assert.Equal(expectedFixedContexts,
            fixedOffsetConverter.ImportContexts.Distinct().OrderBy(item => item.RowIndex).ToArray());
        Assert.Equal(expectedDynamicContexts,
            dynamicValidation.Contexts.Distinct().OrderBy(item => item.RowIndex).ToArray());
        Assert.Equal(expectedRows.Select(row => row + 1),
            fixedValidation.Contexts.Select(context => context.RowIndex).Distinct().OrderBy(row => row));
    }

    /// <summary>
    /// 验证原始日期序列读取器应索引数值单元格并恢复位置。
    /// </summary>
    [Fact]
    public void RawDateSerialReader_ShouldIndexNumericCellsAndRestorePosition()
    {
        using var source = CreateSkippedDateSerialFixture(false, 0, new[] { 61d, 1d, 2d, 3d });
        source.Position = Math.Min(7, source.Length);
        var originalPosition = source.Position;

        var serials = MiniExcelRawDateSerialReader.Read(source, "Data", CancellationToken.None);

        Assert.Equal(12, serials.Count);
        Assert.Equal(61d, serials[MiniExcelRawDateSerialReader.CreateKey(2, 1)]);
        Assert.Equal(1d, serials[MiniExcelRawDateSerialReader.CreateKey(3, 3)]);
        Assert.Equal(2d, serials[MiniExcelRawDateSerialReader.CreateKey(4, 2)]);
        Assert.Equal(3d, serials[MiniExcelRawDateSerialReader.CreateKey(5, 1)]);
        Assert.Equal(originalPosition, source.Position);
    }

    /// <summary>
    /// 测试 - 异步 MiniExcel 导入跳过多行后，日期值和来源行仍与真实工作表位置一致。
    /// </summary>
    /// <param name="date1904">是否启用 1904 日期系统；false 表示 1900 日期系统。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DataRowStartIndexAsync_ShouldAlignRawDateSerials(bool date1904)
    {
        var serials = new[] { 61d, -0.25d, -1.25d, 1d };
        const int headerRowIndex = 2;
        const int skippedDataRows = 2;
        var dataRowStartIndex = headerRowIndex + 1 + skippedDataRows;
        using var source = CreateSkippedDateSerialFixture(date1904, headerRowIndex, serials);
        var fixedOffsetConverter = new FixedOffsetDateConverter();
        var fixedValidation = new FixedColumnContextValidationRule();
        var dynamicValidation = new DynamicColumnContextValidationRule();
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(SkippedDateRow.Offset),
                    ConverterName = fixedOffsetConverter.Name,
                    ValidationRuleNames = { fixedValidation.Name }
                }
            }
        };
        var dynamicDefinition = new ExcelDynamicColumnDefinition
        {
            Key = "dynamicDate",
            Title = "DynamicDate",
            DataType = typeof(DateTime),
            ValidatorName = dynamicValidation.Name
        };
        var request = ExcelImport.Workbook<SkippedDateWorkbook>(workbook =>
            workbook.Sheet<SkippedDateRow>("Data", root => root.Rows, sheet => sheet
                .HeaderRowIndex(headerRowIndex)
                .DataRowStartIndex(dataRowStartIndex)
                .Mapping(mapping)
                .DynamicColumns(row => row.CustomFields, new[] { dynamicDefinition })));
        var expectedDates = GetSkippedDateExpectations(date1904, serials, skippedDataRows);
        var expectedOffsets = expectedDates
            .Select(date => new DateTimeOffset(date, TimeSpan.FromHours(8)))
            .ToArray();
        var expectedRows = Enumerable.Range(dataRowStartIndex, serials.Length - skippedDataRows).ToArray();

        fixedOffsetConverter.ImportContexts.Clear();
        fixedValidation.Contexts.Clear();
        dynamicValidation.Contexts.Clear();
        var result = await new MiniExcelExcelImporter(
            valueConverters: new IExcelValueConverter[] { fixedOffsetConverter },
            namedValidationRules: new INamedExcelValidationRule[] { fixedValidation, dynamicValidation })
            .ImportAsync(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Empty(result.Errors);
        Assert.Equal(expectedDates, result.Workbook.Rows.Select(row => row.Date).ToArray());
        Assert.Equal(expectedOffsets, result.Workbook.Rows.Select(row => row.Offset).ToArray());
        Assert.Equal(expectedDates, result.Workbook.Rows
            .Select(row => (DateTime)row.CustomFields["dynamicDate"]).ToArray());
        Assert.Equal(expectedRows, result.Sheets.Single().SourceRows);
        Assert.Equal(expectedRows.Select(row => row + 1),
            fixedOffsetConverter.ImportContexts.Select(context => context.RowIndex).Distinct().OrderBy(row => row));
        Assert.Equal(expectedRows.Select(row => row + 1),
            dynamicValidation.Contexts.Select(context => context.RowIndex).Distinct().OrderBy(row => row));
    }

    /// <summary>
    /// 验证日期Time偏移导出应保留偏移为不变往返文本。
    /// </summary>
    [Fact]
    public void DateTimeOffsetExport_ShouldPreserveOffsetAsInvariantRoundTripText()
    {
        var expected = new DateTimeOffset(2026, 9, 16, 12, 30, 0, TimeSpan.FromHours(8));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(
            ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
                new[] { new OffsetOnlyRow { OccurredAt = expected } })), source);

        Assert.Contains(expected.ToString("O", CultureInfo.InvariantCulture),
            ReadZipEntry(source, "xl/worksheets/sheet1.xml"));
        source.Position = 0;
        var result = new MiniExcelExcelImporter().Import(source, ExcelImport.Workbook<OffsetOnlyWorkbook>(workbook =>
            workbook.Sheet<OffsetOnlyRow>("Data", root => root.Rows)));

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(expected, Assert.Single(result.Workbook.Rows).OccurredAt);
    }

    /// <summary>
    /// 验证日期Time偏移导出异步应保留偏移用于动态列。
    /// </summary>
    [Fact]
    public async Task DateTimeOffsetExportAsync_ShouldPreserveOffsetForDynamicColumns()
    {
        var expected = new DateTimeOffset(2026, 9, 16, 12, 30, 0, TimeSpan.FromHours(-7));
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "occurredAt",
            Title = "OccurredAt",
            DataType = typeof(DateTimeOffset)
        };
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new DynamicOffsetRow
            {
                CustomFields = new Dictionary<string, object> { ["occurredAt"] = expected }
            } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var source = new MemoryStream();
        await new MiniExcelExcelExporter().ExportAsync(exportRequest, source);

        Assert.Contains(expected.ToString("O", CultureInfo.InvariantCulture),
            ReadZipEntry(source, "xl/worksheets/sheet1.xml"));
        source.Position = 0;
        var result = await new MiniExcelExcelImporter().ImportAsync(source,
            ExcelImport.Workbook<DynamicOffsetWorkbook>(workbook =>
                workbook.Sheet("Data", root => root.Rows,
                    sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition }))));

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(expected, Assert.Single(result.Workbook.Rows).CustomFields["occurredAt"]);
    }

    /// <summary>
    /// 验证读取列范围应报告绝对物理列索引。
    /// </summary>
    [Fact]
    public void ReadColumnRange_ShouldReportAbsolutePhysicalColumnIndex()
    {
        using var source = CreateRangeFixture();
        var request = ExcelImport.Workbook<RangeWorkbook>(workbook =>
            workbook.Sheet<RangeRow>("Data", root => root.Rows, sheet => sheet.ReadColumns(1, 2)));
        source.Position = 0;
        var result = new MiniExcelExcelImporter().Import(source, request);
        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RangeRow.Age), error.PropertyName);
    }

    /// <summary>
    /// 验证关系应遵循自定义键比较器。
    /// </summary>
    [Fact]
    public void Relations_ShouldHonorCustomKeyComparer()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new RelationParent { OrderNo = "A-1" } })
            .AddSheet("Children", new[] { new RelationChild { OrderNo = "a-1", Name = "Item" } }));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<RelationWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });
        var result = new MiniExcelExcelImporter().Import(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Item", Assert.Single(Assert.Single(result.Workbook.Parents).Items).Name);
    }

    /// <summary>
    /// 验证关系导航属性为非 IList 的 ICollection 时仍可完成父子绑定。
    /// </summary>
    [Fact]
    public void Relations_ShouldBindNonListCollectionNavigation()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new CollectionRelationParent { OrderNo = "A-1" } })
            .AddSheet("Children", new[] { new CollectionRelationChild { OrderNo = "a-1", Name = "Item" } }));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<CollectionRelationWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });
        var result = new MiniExcelExcelImporter().Import(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var parent = Assert.Single(result.Workbook.Parents);
        var child = Assert.Single(parent.Items);
        Assert.Equal("Item", child.Name);
        Assert.Equal("a-1", child.OrderNo);
    }

    /// <summary>
    /// 验证关系绑定应保留首个匹配的父级。
    /// </summary>
    [Fact]
    public void Relations_ShouldPreserveFirstMatchingParent()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[]
            {
                new RelationParent { OrderNo = "A-1" },
                new RelationParent { OrderNo = "A-1" }
            })
            .AddSheet("Children", new[]
            {
                new RelationChild { OrderNo = "a-1", Name = "Item" }
            }));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);

        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<RelationWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });
        var result = new MiniExcelExcelImporter().Import(stream, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Parents[0].Items);
        Assert.Empty(result.Workbook.Parents[1].Items);
    }

    /// <summary>
    /// 验证关系没有子级时不会计算父级键。
    /// </summary>
    [Fact]
    public void Relations_WithNoChildren_ShouldNotEvaluateParentKeys()
    {
        var workbook = new RelationWorkbook();
        workbook.Parents.Add(new RelationParent { OrderNo = "unused" });
        var request = ExcelImport.Workbook<RelationWorkbook>(builder =>
        {
            builder.Sheet("Probe", root => root.Parents);
            builder.HasMany(root => root.Parents, root => root.Children,
                _ => ThrowUnusedParentKey(), child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

        var errors = InvokeBindRelations(workbook, request);

        Assert.Empty(errors);
    }

    /// <summary>
    /// 验证关系找到首个匹配父级后会停止评估后续父级。
    /// </summary>
    [Fact]
    public void Relations_ShouldStopAfterFirstMatchingParentBeforeEvaluatingLaterParents()
    {
        var workbook = new RelationWorkbook();
        var first = new RelationParent { OrderNo = "A-1" };
        workbook.Parents.Add(first);
        workbook.Parents.Add(new RelationParent { OrderNo = "unused" });
        var child = new RelationChild { OrderNo = "a-1", Name = "Item" };
        workbook.Children.Add(child);
        var request = ExcelImport.Workbook<RelationWorkbook>(builder =>
        {
            builder.Sheet("Probe", root => root.Parents);
            builder.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo == "A-1" ? parent.OrderNo : ThrowUnusedParentKey(),
                item => item.OrderNo, parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

        var errors = InvokeBindRelations(workbook, request);

        Assert.Empty(errors);
        Assert.Same(child, Assert.Single(first.Items));
        Assert.Empty(workbook.Parents[1].Items);
    }

    /// <summary>
    /// 验证关系会保留重复父级键的评估结果并报告边界错误。
    /// </summary>
    [Fact]
    public void Relations_ShouldPreserveRepeatedParentKeyEvaluationAndErrorBoundary()
    {
        var workbook = new RelationWorkbook();
        workbook.Parents.Add(new RelationParent { OrderNo = "A-1" });
        workbook.Children.Add(new RelationChild { OrderNo = "a-1", Name = "first" });
        workbook.Children.Add(new RelationChild { OrderNo = "a-1", Name = "second" });
        var parentKeyCalls = 0;
        var request = ExcelImport.Workbook<RelationWorkbook>(builder =>
        {
            builder.Sheet("Probe", root => root.Parents);
            builder.HasMany(root => root.Parents, root => root.Children,
                parent =>
                {
                    parentKeyCalls++;
                    if (parentKeyCalls > 1)
                        throw new InvalidOperationException("parent key delegate failed");
                    return parent.OrderNo;
                }, child => child.OrderNo, parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

        var errors = InvokeBindRelations(workbook, request);

        Assert.Equal(2, parentKeyCalls);
        Assert.Equal("first", Assert.Single(workbook.Parents[0].Items).Name);
        var error = Assert.Single(errors);
        Assert.Equal(ExcelImportErrorCode.Relationship, error.Code);
        Assert.Equal("Exception has been thrown by the target of an invocation.", error.Message);
    }

    /// <summary>
    /// 执行关系绑定。
    /// </summary>
    /// <param name="workbook">工作簿对象。</param>
    /// <param name="request">导入或导出请求。</param>
    /// <returns>关系绑定过程中收集的导入错误列表。</returns>
    private static List<ExcelImportError> InvokeBindRelations(RelationWorkbook workbook,
        ExcelWorkbookImportRequest<RelationWorkbook> request)
    {
        var method = typeof(MiniExcelExcelImporter).GetMethod("BindRelations",
            BindingFlags.Static | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(RelationWorkbook));
        var errors = new List<ExcelImportError>();
        method.Invoke(null, new object[] { workbook, request.Relations, errors, request, CancellationToken.None });
        return errors;
    }

    /// <summary>
    /// 验证关系绑定会抛出未使用的父级键错误。
    /// </summary>
    /// <returns>此方法始终抛出异常，不返回键值。</returns>
    private static string ThrowUnusedParentKey() =>
        throw new InvalidOperationException("unused parent key evaluated");

    /// <summary>
    /// 验证资源限制失败时会在 MiniExcel 解析器执行前终止。
    /// </summary>
    [Fact]
    public void ResourceLimit_ShouldFailBeforeMiniExcelParser()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new Person { Name = "limited", Age = 1 } }));
        using var stream = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, stream);
        stream.Position = 0;
        var importRequest = ExcelImport.Workbook<Workbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
            .Sheet<Person>("Data", root => root.People));

        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new MiniExcelExcelImporter().Import(stream, importRequest));
    }

    /// <summary>
    /// 验证失败工作簿会在解析器执行前被拒绝。
    /// </summary>
    [Fact]
    public void FailureWorkbook_ShouldBeRejectedBeforeParser()
    {
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new Person { Name = "failure", Age = 1 } }));
        using var source = new MemoryStream();
        new MiniExcelExcelExporter().Export(exportRequest, source);
        source.Position = 0;
        using var failureDestination = new MemoryStream();
        var importRequest = ExcelImport.Workbook<Workbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureDestination
            })
            .Sheet<Person>("Data", root => root.People));

        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelImporter().Import(source, importRequest));
        Assert.Equal(0, failureDestination.Length);
    }

    /// <summary>
    /// 验证预先取消的异步操作不会写入或读取。
    /// </summary>
    [Fact]
    public async Task PreCanceledAsync_ShouldNotWriteOrRead()
    {
        using var stream = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var exportRequest = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new Person { Name = "cancel", Age = 1 } }));

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new MiniExcelExcelExporter().ExportAsync(exportRequest, stream, cancellation.Token));
        Assert.Equal(0, stream.Length);
    }

    /// <summary>
    /// 验证异步导出写入期间取消时会保持目标打开。
    /// </summary>
    [Fact]
    public async Task AsyncExport_MidWriteCancellation_ShouldKeepDestinationOpen()
    {
        using var cancellation = new CancellationTokenSource();
        using var destination = new CancelOnFirstAsyncWriteStream(cancellation);
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            Enumerable.Range(0, 100).Select(index => new Person { Name = $"row-{index}", Age = index })));

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new MiniExcelExcelExporter().ExportAsync(request, destination, cancellation.Token));

        Assert.True(destination.CanWrite);
        Assert.True(destination.AsyncWriteCount > 0);
    }

    /// <summary>
    /// 验证异步导入读取期间取消时会保持源打开。
    /// </summary>
    [Fact]
    public async Task AsyncImport_MidReadCancellation_ShouldKeepSourceOpen()
    {
        using var sourceBytes = new MemoryStream();
        new MiniExcelExcelExporter().Export(
            ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
                new[] { new Person { Name = "read", Age = 1 } })), sourceBytes);
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelOnFirstAsyncReadStream(sourceBytes.ToArray(), cancellation);
        var request = ExcelImport.Workbook<Workbook>(workbook => workbook
            .Sheet<Person>("Data", root => root.People));

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new MiniExcelExcelImporter().ImportAsync(source, request, cancellation.Token));

        Assert.True(source.CanRead);
        Assert.True(source.AsyncReadCount > 0);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
    }

    /// <summary>
    /// 验证异步文件导出预先取消时会保留现有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task AsyncFileExport_PreCanceled_ShouldPreserveExistingTargetAndCleanTempFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "target.xlsx");
        var original = Encoding.UTF8.GetBytes("before");
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
                new[] { new Person { Name = "cancel", Age = 1 } }));
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new MiniExcelExcelExporter().ExportToFileAsync(request, path, cancellation.Token));

            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "target.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证异步文件导出期间取消时会保留现有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTempFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "target.xlsx");
        var original = Encoding.UTF8.GetBytes("before-mid-flight");
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstExportConverter(cancellation);
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
                Enumerable.Range(0, 1000).Select(index => new MidFlightRow
                {
                    Code = $"row-{index}"
                }), sheet => sheet.Mapping(new ExcelMappingConfiguration
                {
                    Columns =
                    {
                        new ExcelColumnConfiguration
                        {
                            PropertyName = nameof(MidFlightRow.Code),
                            Title = "Code",
                            ConverterName = "cancel-after-first"
                        }
                    }
                })));

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new MiniExcelExcelExporter(new IExcelValueConverter[] { converter })
                    .ExportToFileAsync(request, path, cancellation.Token));

            Assert.True(converter.ExportCalls > 0);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "target.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class Workbook
    {
        /// <summary>
        /// 获取人员集合。
        /// </summary>
        public List<Person> People { get; } = new();
        /// <summary>
        /// 获取团队集合。
        /// </summary>
        public List<Team> Teams { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的人员数据。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class DateWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<DateRow> Rows { get; } = new();
    }

    /// <summary>
    /// 跳过行日期合同测试的工作簿根实体。
    /// </summary>
    private sealed class SkippedDateWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<SkippedDateRow> Rows { get; } = new();
    }

    /// <summary>
    /// 同时包含固定日期、固定 offset 日期和动态日期列的测试行实体。
    /// </summary>
    private sealed class SkippedDateRow
    {
        /// <summary>
        /// 获取或设置日期。
        /// </summary>
        [ExcelDate]
        public DateTime Date { get; set; }

        /// <summary>
        /// 获取或设置偏移量。
        /// </summary>
        public DateTimeOffset Offset { get; set; }

        /// <summary>
        /// 获取或设置自定义字段集合。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> CustomFields { get; set; } = new();
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class FixedOffsetDateConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "fixed-offset-date";

        /// <summary>
        /// 获取记录转换器收到的实际一基行列上下文，供物理位置合同断言使用。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> ImportContexts { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(DateTimeOffset);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportContexts.Add((context.RowIndex, context.ColumnIndex));
            var attribute = new ExcelDateAttribute
            {
                OffsetPolicy = ExcelDateOffsetPolicy.UseFixedOffset,
                OffsetMinutes = 480
            };
            return new DateTimeExcelValidationRule().TryParseValue(context.Cell,
                context.Value?.ToString() ?? string.Empty, context.PropertyType, context.Culture,
                attribute, out value);
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class DateRow
    {
        /// <summary>
        /// 获取或设置日期。
        /// </summary>
        [ExcelDate]
        public DateTime Date { get; set; }

        /// <summary>
        /// 获取或设置偏移量。
        /// </summary>
        public DateTimeOffset Offset { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class OffsetOnlyWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<OffsetOnlyRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class OffsetOnlyRow
    {
        /// <summary>
        /// 获取或设置发生时间。
        /// </summary>
        public DateTimeOffset OccurredAt { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class DynamicOffsetWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<DynamicOffsetRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class DynamicOffsetRow
    {
        /// <summary>
        /// 获取或设置自定义字段集合。
        /// </summary>
        public Dictionary<string, object> CustomFields { get; set; } = new();
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class RangeWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<RangeRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class RangeRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 表示关系测试使用的团队数据。
    /// </summary>
    private sealed class Team
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class DynamicWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<DynamicRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置自定义字段集合。
        /// </summary>
        public Dictionary<string, object> CustomFields { get; set; } = new();
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class DynamicContextWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<DynamicContextRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class DynamicContextRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public int Code { get; set; }
        /// <summary>
        /// 获取或设置自定义字段集合。
        /// </summary>
        public Dictionary<string, object> CustomFields { get; set; } = new();
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class FixedColumnContextWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<FixedColumnContextRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class FixedColumnContextSourceRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public string Age { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class FixedColumnContextRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ValidationModeWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ValidationModeRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ValidationModeRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class ValueMapWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<ValueMapRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ValueMapRow
    {
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        [ValueMapping("Enabled", 1)]
        public int Status { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class UniqueWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<UniqueRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class UniqueRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelUnique(IgnoreEmpty = false)]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
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
    /// 表示关系映射测试中的父项数据。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取明细项集合。
        /// </summary>
        public List<RelationChild> Items { get; } = new();
    }

    /// <summary>
    /// 表示关系映射测试中的子项数据。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示非 IList 关系测试使用的工作簿数据模型。
    /// </summary>
    private sealed class CollectionRelationWorkbook
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<CollectionRelationParent> Parents { get; } = new();
        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<CollectionRelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示非 IList 关系测试中的父项数据。
    /// </summary>
    private sealed class CollectionRelationParent
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取由 HashSet 实现的导航集合。
        /// </summary>
        public ICollection<CollectionRelationChild> Items { get; } = new HashSet<CollectionRelationChild>();
    }

    /// <summary>
    /// 表示非 IList 关系测试中的子项数据。
    /// </summary>
    private sealed class CollectionRelationChild
    {
        /// <summary>
        /// 获取或设置订单号。
        /// </summary>
        public string OrderNo { get; set; }
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class UpperRegionConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "upper-region";
        /// <summary>
        /// 获取已导出行集合。
        /// </summary>
        public List<int> ExportRows { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value?.ToString()?.ToUpperInvariant();
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ExportRows.Add(context.RowIndex);
            value = context.Value?.ToString()?.ToUpperInvariant();
            return true;
        }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class RowIndexConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "row-index";
        /// <summary>
        /// 获取已导出行集合。
        /// </summary>
        public List<int> ExportRows { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ExportRows.Add(context.RowIndex);
            value = context.Value;
            return true;
        }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class CancelAfterFirstExportConverter : INamedExcelValueConverter
    {
        /// <summary>
        /// 用于在首次导出后取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 指示是否已触发首次导出后的取消。
        /// </summary>
        private bool _cancelled;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstExportConverter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">首次导出转换时触发的取消源。</param>
        public CancelAfterFirstExportConverter(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <inheritdoc />
        public string Name => "cancel-after-first";
        /// <summary>
        /// 获取或设置导出调用次数。
        /// </summary>
        public int ExportCalls { get; private set; }

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ExportCalls++;
            value = context.Value;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }
            return true;
        }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class FixedColumnContextConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "capture-fixed-context";
        /// <summary>
        /// 获取导入上下文集合。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> ImportContexts { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportContexts.Add((context.RowIndex, context.ColumnIndex));
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }
    }

    /// <summary>
    /// 提供测试场景使用的校验规则。
    /// </summary>
    private sealed class FixedColumnContextValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "capture-fixed-context-validation";
        /// <inheritdoc />
        public string ErrorMessage => "固定列值无效。";
        /// <summary>
        /// 获取上下文集合。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> Contexts { get; } = new();

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context)
        {
            Contexts.Add((context.RowIndex, context.ColumnIndex));
            return true;
        }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class MidFlightRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class DynamicColumnContextConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "capture-dynamic-context";
        /// <summary>
        /// 获取导入上下文集合。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> ImportContexts { get; } = new();

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportContexts.Add((context.RowIndex, context.ColumnIndex));
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }
    }

    /// <summary>
    /// 提供测试场景使用的校验规则。
    /// </summary>
    private sealed class DynamicColumnContextValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "reject-invalid-dynamic";
        /// <inheritdoc />
        public string ErrorMessage => "动态列值无效。";
        /// <summary>
        /// 获取上下文集合。
        /// </summary>
        public List<(int RowIndex, int ColumnIndex)> Contexts { get; } = new();

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context)
        {
            Contexts.Add((context.RowIndex, context.ColumnIndex));
            return !string.Equals(context.Value, "invalid", StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class CancelOnFirstAsyncWriteStream : Stream
    {
        /// <summary>
        /// 承载异步写入测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 用于在首次异步写入后取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 指示是否已触发首次异步写入后的取消。
        /// </summary>
        private bool _cancelled;

        /// <summary>
        /// 初始化一个 <see cref="CancelOnFirstAsyncWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">首次异步写入后触发的取消源。</param>
        public CancelOnFirstAsyncWriteStream(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <summary>
        /// 获取或设置异步写入次数。
        /// </summary>
        public int AsyncWriteCount { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => true;
        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            AsyncWriteCount++;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }
            return _inner.WriteAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }
            return _inner.WriteAsync(buffer, cancellationToken);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供测试场景使用的流替身。
    /// </summary>
    private sealed class CancelOnFirstAsyncReadStream : Stream
    {
        /// <summary>
        /// 承载异步读取测试数据的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 用于在首次异步读取后取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 指示是否已触发首次异步读取后的取消。
        /// </summary>
        private bool _cancelled;

        /// <summary>
        /// 初始化一个 <see cref="CancelOnFirstAsyncReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="bytes">待异步读取的输入字节。</param>
        /// <param name="cancellation">首次异步读取后触发的取消源。</param>
        public CancelOnFirstAsyncReadStream(byte[] bytes, CancellationTokenSource cancellation)
        {
            _inner = new MemoryStream(bytes, writable: false);
            _cancellation = cancellation;
        }

        /// <summary>
        /// 获取或设置异步读取次数。
        /// </summary>
        public int AsyncReadCount { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;
        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Task.CompletedTask;
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            AsyncReadCount++;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }
            return _inner.ReadAsync(buffer, offset, count, cancellationToken);
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncReadCount++;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }
            return _inner.ReadAsync(buffer, cancellationToken);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 将工作簿标记为 1904 日期系统。
    /// </summary>
    /// <param name="stream">参与操作的流。</param>
    private static void MarkWorkbookAsDate1904(MemoryStream stream)
    {
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            Assert.NotNull(entry);
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true))
                xml = reader.ReadToEnd();
            entry.Delete();
            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var root = document.Root;
            Assert.NotNull(root);
            var workbookPr = root.Element(root.Name.Namespace + "workbookPr");
            if (workbookPr == null)
            {
                workbookPr = new XElement(root.Name.Namespace + "workbookPr");
                root.AddFirst(workbookPr);
            }
            workbookPr.SetAttributeValue("date1904", "1");
            xml = document.ToString(SaveOptions.DisableFormatting);
            using var writer = new StreamWriter(archive.CreateEntry("xl/workbook.xml").Open(),
                new UTF8Encoding(false));
            writer.Write(xml);
        }
        stream.Position = 0;
    }

    /// <summary>
    /// 创建日期序列测试夹具。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <returns>包含日期序列值且位置归零的 XLSX 内存流，由调用方释放。</returns>
    private static MemoryStream CreateDateSerialFixture(bool date1904)
    {
        var stream = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Date");
            header.CreateCell(1).SetCellValue("Offset");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
            var serials = new[] { -2.25d, -1.25d, -1d, -0.25d, 0d, 1d, 59d, 60d, 61d };
            for (var index = 0; index < serials.Length; index++)
            {
                var row = sheet.CreateRow(index + 1);
                var cell = row.CreateCell(0);
                cell.SetCellValue(serials[index]);
                cell.CellStyle = style;
                var offsetCell = row.CreateCell(1);
                offsetCell.SetCellValue(serials[index]);
                offsetCell.CellStyle = style;
            }
            workbook.Write(stream, true);
        }
        MarkWorkbookDateSystem(stream, date1904);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建带正文间隔的真实 XLSX 日期 serial fixture。
    /// </summary>
    /// <param name="date1904">是否设置工作簿的 1904 日期系统标志。</param>
    /// <param name="headerRowIndex">表头所在的零基行索引。</param>
    /// <param name="serials">按连续正文行写入的原始日期 serial。</param>
    /// <returns>定位在开头且保持可读写的 XLSX 内存流。</returns>
    private static MemoryStream CreateSkippedDateSerialFixture(bool date1904, int headerRowIndex,
        IReadOnlyList<double> serials)
    {
        var stream = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(headerRowIndex);
            header.CreateCell(0).SetCellValue("Date");
            header.CreateCell(1).SetCellValue("Offset");
            header.CreateCell(2).SetCellValue("DynamicDate");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm");
            for (var index = 0; index < serials.Count; index++)
            {
                var row = sheet.CreateRow(headerRowIndex + index + 1);
                for (var column = 0; column < 3; column++)
                {
                    var cell = row.CreateCell(column);
                    cell.SetCellValue(serials[index]);
                    cell.CellStyle = style;
                }
            }
            workbook.Write(stream, true);
        }
        MarkWorkbookDateSystem(stream, date1904);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 根据独立的 1900/1904 合同表计算跳过行测试的完整日期预期。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <param name="serials">fixture 中按物理行排列的 serial。</param>
    /// <param name="skippedDataRows">需要跳过的正文行数量。</param>
    /// <returns>从正文起始行开始的完整日期预期。</returns>
    private static DateTime[] GetSkippedDateExpectations(bool date1904, IReadOnlyList<double> serials,
        int skippedDataRows)
    {
        var expected = date1904
            ? new Dictionary<double, DateTime>
            {
                [61d] = new DateTime(1904, 3, 2),
                [-0.25d] = new DateTime(1903, 12, 31, 18, 0, 0),
                [-1.25d] = new DateTime(1903, 12, 30, 18, 0, 0),
                [1d] = new DateTime(1904, 1, 2)
            }
            : new Dictionary<double, DateTime>
            {
                [61d] = new DateTime(1900, 3, 1),
                [-0.25d] = new DateTime(1899, 12, 30, 18, 0, 0),
                [-1.25d] = new DateTime(1899, 12, 29, 18, 0, 0),
                [1d] = new DateTime(1900, 1, 1)
            };
        return serials.Skip(skippedDataRows).Select(serial => expected[serial]).ToArray();
    }

    /// <summary>
    /// 创建范围测试夹具。
    /// </summary>
    /// <returns>包含有效名称和无效整数值且位置归零的 XLSX 内存流。</returns>
    private static MemoryStream CreateRangeFixture()
    {
        var stream = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Unused");
            header.CreateCell(1).SetCellValue("Name");
            header.CreateCell(2).SetCellValue("Age");
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("unused");
            row.CreateCell(1).SetCellValue("valid");
            row.CreateCell(2).SetCellValue("not-an-int");
            workbook.Write(stream, true);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 读取 ZIP 条目文本。
    /// </summary>
    /// <param name="source">输入流。</param>
    /// <param name="entryName">ZIP 条目名称。</param>
    /// <returns>指定 ZIP 条目的完整 UTF-8 文本。</returns>
    private static string ReadZipEntry(Stream source, string entryName)
    {
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var reader = new StreamReader(archive.GetEntry(entryName).Open(), Encoding.UTF8, true);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// 设置工作簿日期系统标志。
    /// </summary>
    /// <param name="stream">参与操作的流。</param>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    private static void MarkWorkbookDateSystem(MemoryStream stream, bool date1904)
    {
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            XDocument document;
            using (var input = entry.Open())
                document = XDocument.Load(input);
            var ns = document.Root.Name.Namespace;
            var properties = document.Root.Element(ns + "workbookPr");
            if (properties == null)
            {
                properties = new XElement(ns + "workbookPr");
                document.Root.AddFirst(properties);
            }
            properties.SetAttributeValue("date1904", date1904 ? "1" : "0");
            entry.Delete();
            using var output = archive.CreateEntry("xl/workbook.xml").Open();
            document.Save(output);
        }
        stream.Position = 0;
    }
}
