using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Npoi.Tests;

/// <summary>
/// NPOI Provider 的日期系统、物理位置和结构化错误回归。
/// </summary>
public sealed class NpoiProviderBoundaryRegressionTest
{
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
            new DateTime(1900, 2, 28), new DateTime(1900, 2, 28), new DateTime(1900, 3, 1)
        };
        var expected1904 = new[]
        {
            new DateTime(1903, 12, 29, 18, 0, 0), new DateTime(1903, 12, 30, 18, 0, 0),
            new DateTime(1903, 12, 31), new DateTime(1903, 12, 31, 18, 0, 0),
            new DateTime(1904, 1, 1), new DateTime(1904, 1, 2), new DateTime(1904, 2, 29),
            new DateTime(1904, 3, 1), new DateTime(1904, 3, 2)
        };

        foreach (var date1904 in new[] { false, true })
        {
            using var source = CreateDateSerialFixture(date1904);
            var converter = new FixedOffsetDateConverter();
            var request = ExcelImport.Workbook<DateWorkbook>(workbook =>
                workbook.Sheet<DateRow>("Data", root => root.Rows, sheet => sheet.Mapping(
                    new ExcelMappingConfiguration
                    {
                        Columns =
                        {
                            new ExcelColumnConfiguration
                            {
                                PropertyName = nameof(DateRow.Offset),
                                ConverterName = converter.Name
                            }
                        }
                    })));
            var expectedDates = date1904 ? expected1904 : expected1900;
            var expectedOffsets = expectedDates
                .Select(date => new DateTimeOffset(date, TimeSpan.FromHours(8))).ToArray();

            var result = new NpoiExcelImporter(valueConverters: new[] { converter }).Import(source, request);

            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(expectedDates, result.Workbook.Rows.Select(row => row.Date).ToArray());
            Assert.Equal(expectedOffsets, result.Workbook.Rows.Select(row => row.Offset).ToArray());
        }
    }

    /// <summary>
    /// 验证数据行开始索引与原始日期序列保持一致。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <param name="headerRowIndex">数据行或行集合。</param>
    /// <param name="skippedDataRows">数据行或行集合。</param>
    [Theory]
    [InlineData(false, 0, 1)]
    [InlineData(true, 0, 1)]
    [InlineData(false, 2, 2)]
    [InlineData(true, 2, 2)]
    public void DataRowStartIndex_ShouldAlignRawDateSerials(bool date1904, int headerRowIndex,
        int skippedDataRows)
    {
        var serials = new[] { 61d, -0.25d, -1.25d, 1d };
        var dataRowStartIndex = headerRowIndex + 1 + skippedDataRows;
        using var source = CreateSkippedDateSerialFixture(date1904, headerRowIndex, serials);
        var converter = new FixedOffsetDateConverter();
        var fixedValidation = new FixedColumnContextValidationRule();
        var dynamicValidation = new DynamicColumnContextValidationRule();
        var request = ExcelImport.Workbook<SkippedDateWorkbook>(workbook =>
            workbook.Sheet<SkippedDateRow>("Data", root => root.Rows, sheet => sheet
                .HeaderRowIndex(headerRowIndex)
                .DataRowStartIndex(dataRowStartIndex)
                .Mapping(new ExcelMappingConfiguration
                {
                    Columns =
                    {
                        new ExcelColumnConfiguration
                        {
                            PropertyName = nameof(SkippedDateRow.Offset),
                            ConverterName = converter.Name,
                            ValidationRuleNames = { fixedValidation.Name }
                        }
                    }
                })
                .DynamicColumns(row => row.CustomFields, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "dynamicDate", Title = "DynamicDate", DataType = typeof(DateTime),
                        ValidatorName = dynamicValidation.Name
                    }
                })));
        var expectedDates = GetSkippedDateExpectations(date1904, serials, skippedDataRows);
        var expectedRows = Enumerable.Range(dataRowStartIndex, serials.Length - skippedDataRows).ToArray();
        var expectedOffsets = expectedDates.Select(date => new DateTimeOffset(date, TimeSpan.FromHours(8))).ToArray();

        var result = new NpoiExcelImporter(valueConverters: new[] { converter },
            namedValidationRules: new INamedExcelValidationRule[] { fixedValidation, dynamicValidation })
            .Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Empty(result.Errors);
        Assert.Equal(expectedDates, result.Workbook.Rows.Select(row => row.Date).ToArray());
        Assert.Equal(expectedOffsets, result.Workbook.Rows.Select(row => row.Offset).ToArray());
        Assert.Equal(expectedDates, result.Workbook.Rows
            .Select(row => (DateTime)row.CustomFields["dynamicDate"]).ToArray());
        Assert.Equal(expectedRows, result.Sheets.Single().SourceRows);
        Assert.Equal(expectedRows.Select(row => (RowIndex: row + 1, ColumnIndex: 2)),
            converter.ImportContexts.Distinct().OrderBy(item => item.RowIndex).ToArray());
        Assert.Equal(expectedRows.Select(row => (RowIndex: row + 1, ColumnIndex: 2)),
            fixedValidation.Contexts.Distinct().OrderBy(item => item.RowIndex).ToArray());
        Assert.Equal(expectedRows.Select(row => (RowIndex: row + 1, ColumnIndex: 3)),
            dynamicValidation.Contexts.Distinct().OrderBy(item => item.RowIndex).ToArray());
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

        var result = new NpoiExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RangeRow.Age), error.PropertyName);
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
    private sealed class SkippedDateWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<SkippedDateRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
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
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class FixedOffsetDateConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "npoi-fixed-offset-date";
        /// <summary>
        /// 获取导入上下文集合。
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
    /// 提供测试场景使用的校验规则。
    /// </summary>
    private sealed class FixedColumnContextValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "npoi-fixed-context";
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
    /// 提供测试场景使用的校验规则。
    /// </summary>
    private sealed class DynamicColumnContextValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "npoi-dynamic-context";
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
            return true;
        }
    }

    /// <summary>
    /// 创建日期序列测试夹具。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <returns>生成的流。</returns>
    private static MemoryStream CreateDateSerialFixture(bool date1904)
    {
        var stream = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Date");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Offset");
            var style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
            var serials = new[] { -2.25d, -1.25d, -1d, -0.25d, 0d, 1d, 59d, 60d, 61d };
            for (var index = 0; index < serials.Length; index++)
            {
                var row = sheet.CreateRow(index + 1);
                row.CreateCell(0).SetCellValue(serials[index]);
                row.GetCell(0).CellStyle = style;
                row.CreateCell(1).SetCellValue(serials[index]);
                row.GetCell(1).CellStyle = style;
            }
            workbook.Write(stream, true);
        }
        MarkWorkbookDateSystem(stream, date1904);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建跳过日期序列的测试夹具。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <param name="headerRowIndex">数据行或行集合。</param>
    /// <param name="serials">日期序列值。</param>
    /// <returns>生成的流。</returns>
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
                    row.CreateCell(column).SetCellValue(serials[index]);
                    row.GetCell(column).CellStyle = style;
                }
            }
            workbook.Write(stream, true);
        }
        MarkWorkbookDateSystem(stream, date1904);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建范围测试夹具。
    /// </summary>
    /// <returns>生成的流。</returns>
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
    /// 获取跳过行的日期预期值。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <param name="serials">日期序列值。</param>
    /// <param name="skippedDataRows">数据行或行集合。</param>
    /// <returns>生成的集合。</returns>
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
            Assert.NotNull(entry);
            XDocument document;
            using (var input = entry.Open())
                document = XDocument.Load(input);
            var root = document.Root;
            Assert.NotNull(root);
            var ns = root.Name.Namespace;
            var properties = root.Element(ns + "workbookPr") ?? new XElement(ns + "workbookPr");
            properties.SetAttributeValue("date1904", date1904 ? "1" : "0");
            if (properties.Parent == null)
                root.AddFirst(properties);
            entry.Delete();
            using var output = archive.CreateEntry("xl/workbook.xml").Open();
            document.Save(output);
        }
        stream.Position = 0;
    }
}
