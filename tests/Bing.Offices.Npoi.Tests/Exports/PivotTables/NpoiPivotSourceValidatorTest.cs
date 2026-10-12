using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 高级透视来源计划的类型、坐标、筛选和展开边界。
/// </summary>
public sealed class NpoiPivotSourceValidatorTest
{
    private static readonly string[] Headers =
        { "Region", "Amount", "Status", "Numeric", "Active", "When", "Optional", "Unused" };

    private static readonly object[][] BasicRows =
    {
        new object[] { "North", 10d, "Retail", 12d, false, new DateTime(2025, 1, 2), null, "a" },
        new object[] { "South", 20d, "RETAIL", 15d, true, new DateTime(2025, 1, 3), "filled", "b" },
        new object[] { "North", 30d, "Retail ", 12d, false, new DateTime(2025, 1, 2), string.Empty, "c" }
    };

    /// <summary>
    /// 验证偏移区域、大小写筛选及计划快照保留全部真实来源值。
    /// </summary>
    [Fact]
    public void OffsetSource_ShouldPreserveRowsAndResolveTypedFilterValues()
    {
        using var workbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[] { Value("Amount", "Total") },
            filters: new[]
            {
                Filter("Status", "retail"), Filter("Numeric", 12), Filter("Active", false),
                Filter("When", new DateTime(2025, 1, 2)), Filter("Optional", null)
            });

        var plan = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);

        Assert.Equal(Headers, plan.Headers);
        Assert.Equal(3, plan.Rows.Count);
        for (var index = 0; index < BasicRows.Length; index++)
            Assert.Equal(BasicRows[index], plan.Rows[index]);
        Assert.Equal(0, plan.Columns["Region"]);
        Assert.Equal(1, plan.Columns["Amount"]);
        Assert.Equal("Retail", Assert.IsType<string>(Assert.Single(plan.Filters["Status"])));
        Assert.Equal(12d, Assert.IsType<double>(Assert.Single(plan.Filters["Numeric"])));
        Assert.False(Assert.IsType<bool>(Assert.Single(plan.Filters["Active"])));
        Assert.Equal(new DateTime(2025, 1, 2), Assert.IsType<DateTime>(Assert.Single(plan.Filters["When"])));
        Assert.Null(Assert.Single(plan.Filters["Optional"]));
        Assert.Throws<NotSupportedException>(() => ((IList<IReadOnlyList<object>>)plan.Rows).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<object>)plan.Rows[0])[0] = "changed");
    }

    /// <summary>
    /// 验证筛选不会 trim 文本，且不匹配的文本数值保持类型区别。
    /// </summary>
    [Fact]
    public void FilterMatching_ShouldIgnoreCaseWithoutTrimmingOrConvertingTypes()
    {
        using var workbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2);
        var range = new ExcelRangeDefinition { StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9 };
        var padded = AdvancedDefinition(range, new[] { Axis("Region") }, new[] { Value("Amount", "Total") },
            new[] { Filter("Status", "Retail ") });
        var plan = Validate(workbook, padded);

        Assert.Equal("Retail ", Assert.IsType<string>(Assert.Single(plan.Filters["Status"])));
        Assert.Equal(3, plan.Rows.Count);

        var wrongNumericType = AdvancedDefinition(range, new[] { Axis("Region") },
            new[] { Value("Amount", "Total") }, new[] { Filter("Numeric", "12") });
        var exception = CaptureLocation(workbook, wrongNumericType);
        Assert.Equal(4, exception.RowIndex);
        Assert.Equal(6, exception.ColumnIndex);
        Assert.Contains("指定值不在来源中", exception.Message);
    }

    /// <summary>
    /// 验证表头唯一性和缺失字段错误指向原始偏移坐标。
    /// </summary>
    [Fact]
    public void InvalidHeadersAndMissingFields_ShouldReportOriginalCoordinates()
    {
        using (var blankWorkbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2))
        {
            blankWorkbook.GetSheet("Data").GetRow(3).GetCell(4).SetCellValue("  ");
            var blank = CaptureLocation(blankWorkbook, LegacyDefinition(new ExcelRangeDefinition
                {
                    StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
                }));
            AssertLocation(blank, 4, 5);
        }

        using (var duplicateWorkbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2))
        {
            duplicateWorkbook.GetSheet("Data").GetRow(3).GetCell(4).SetCellValue("region");
            var exception = CaptureLocation(duplicateWorkbook, LegacyDefinition(new ExcelRangeDefinition
            {
                StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
            }));
            AssertLocation(exception, 4, 5);
        }

        using var missingWorkbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2);
        var missing = CaptureLocation(missingWorkbook, LegacyDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
        }, rowField: "NotThere"));
        AssertLocation(missing, 4, 3);
        Assert.Contains("NotThere", missing.Message);
    }

    /// <summary>
    /// 验证未参与字段角色的公式或错误单元格也会在来源读取时拒绝。
    /// </summary>
    /// <param name="errorCell">true 表示写入 Excel 错误，false 表示写入公式。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormulaOrErrorInUnusedColumn_ShouldReportItsOriginalCoordinate(bool errorCell)
    {
        using var workbook = CreateWorkbook(BasicRows, Headers, startRow: 3, startColumn: 2);
        var cell = workbook.GetSheet("Data").GetRow(4).GetCell(9);
        if (errorCell)
            cell.SetCellErrorValue(FormulaError.VALUE.Code);
        else
            cell.SetCellFormula("1+2");
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[] { Value("Amount", "Total") });

        var exception = CaptureLocation(workbook, definition);

        Assert.Equal(5, exception.RowIndex);
        Assert.Equal(10, exception.ColumnIndex);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains(errorCell ? "错误值" : "公式", exception.Message);
    }

    /// <summary>
    /// 验证数值聚合拒绝文本、日期、布尔及全空度量并指向首个非法值。
    /// </summary>
    /// <param name="kind">待测度量值类别。</param>
    [Theory]
    [InlineData("text", ExcelPivotAggregation.Sum)]
    [InlineData("date", ExcelPivotAggregation.Average)]
    [InlineData("boolean", ExcelPivotAggregation.Minimum)]
    [InlineData("all-empty", ExcelPivotAggregation.Maximum)]
    public void NumericAggregation_ShouldRejectNonNumericOrAllEmpty(string kind, ExcelPivotAggregation aggregation)
    {
        var rows = CloneRows(BasicRows);
        for (var index = 0; index < rows.Length; index++)
        {
            rows[index][1] = kind switch
            {
                "text" => "12",
                "date" => new DateTime(2025, 1, index + 1),
                "boolean" => index % 2 == 0,
                _ => null
            };
        }
        using var workbook = CreateWorkbook(rows, Headers, startRow: 3, startColumn: 2);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", Aggregation = aggregation }
        });

        var exception = CaptureLocation(workbook, definition);

        Assert.Equal(5, exception.RowIndex);
        Assert.Equal(4, exception.ColumnIndex);
        Assert.Contains(kind == "all-empty" ? "没有数值" : "必须为数值", exception.Message);
    }

    /// <summary>
    /// 验证完全空白的数据行被拒绝并报告来源区域中的原始位置。
    /// </summary>
    [Fact]
    public void EmptySourceRow_ShouldReportOriginalRowAndColumn()
    {
        var rows = CloneRows(BasicRows);
        rows[1] = new object[Headers.Length];
        using var workbook = CreateWorkbook(rows, Headers, startRow: 3, startColumn: 2);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 6, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[] { Value("Amount", "Total") });

        var exception = CaptureLocation(workbook, definition);

        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(3, exception.ColumnIndex);
        Assert.Contains("空数据行", exception.Message);
    }

    /// <summary>
    /// 验证 Count 保留混合来源类型并生成完整透视缓存。
    /// </summary>
    [Fact]
    public void CountAggregation_ShouldAcceptMixedTypedAndEmptyValues()
    {
        var rows = new[]
        {
            new object[] { "North", 10d, "Retail", 1d, true, new DateTime(2025, 1, 2), "text", "x" },
            new object[] { "South", 20d, "Retail", 2d, false, new DateTime(2025, 1, 3), 42d, "x" },
            new object[] { "West", 30d, "Retail", 3d, true, new DateTime(2025, 1, 4), false, "x" },
            new object[] { "East", 40d, "Retail", 4d, false, new DateTime(2025, 1, 5), new DateTime(2025, 1, 6), "x" },
            new object[] { "Blank", 50d, "Retail", 5d, true, new DateTime(2025, 1, 7), string.Empty, "x" },
            new object[] { "Missing", 60d, "Retail", 6d, false, new DateTime(2025, 1, 8), null, "x" }
        };
        using var workbook = CreateWorkbook(rows, Headers, startRow: 3, startColumn: 2);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 9, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[]
        {
            new ExcelPivotValueDefinition
                { Field = "Optional", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
        });

        var plan = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);

        Assert.Equal(6, plan.Rows.Count);
        Assert.Equal(new[] { "text", "double", "bool", "DateTime", "string.Empty", "null" },
            plan.Rows.Select(row => row[6] switch
            {
                null => "null",
                string text when text.Length == 0 => "string.Empty",
                string => "text",
                double => "double",
                bool => "bool",
                DateTime => "DateTime",
                _ => "unexpected"
            }));

        var records = NpoiPivotTableWriter.Create(workbook, plan, CancellationToken.None);
        using var output = new MemoryStream();
        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument ReadPart(string prefix)
        {
            var entry = Assert.Single(archive.Entries.Where(item =>
                item.FullName.Contains(prefix, StringComparison.Ordinal)
                && item.FullName.EndsWith(".xml", StringComparison.Ordinal)));
            using var stream = entry.Open();
            return XDocument.Load(stream);
        }

        var pivot = ReadPart("pivotTable").Root;
        var countField = Assert.Single(pivot.Element(spreadsheet + "dataFields")
            .Elements(spreadsheet + "dataField"));
        Assert.Equal("count", (string)countField.Attribute("subtotal"));

        var cache = ReadPart("pivotCacheDefinition").Root;
        Assert.Equal((uint)rows.Length, (uint)cache.Attribute("recordCount"));
        var fields = cache.Element(spreadsheet + "cacheFields").Elements(spreadsheet + "cacheField").ToArray();
        var optionalField = fields[6];
        var sharedItems = optionalField.Element(spreadsheet + "sharedItems").Elements().ToArray();
        Assert.Equal(new[] { "s", "n", "b", "d", "s", "m" },
            sharedItems.Select(item => item.Name.LocalName));
        Assert.Equal(string.Empty, (string)sharedItems[4].Attribute("v"));

        var cacheRows = ReadPart("pivotCacheRecords").Root.Elements(spreadsheet + "r").ToArray();
        Assert.Equal(rows.Length, cacheRows.Length);
        Assert.Equal(new[] { "s", "n", "b", "d", "s", "m" }, cacheRows.Select(row =>
        {
            var index = (int)row.Elements(spreadsheet + "x").ElementAt(6).Attribute("v");
            return sharedItems[index].Name.LocalName;
        }));

        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        Assert.Single(((XSSFSheet)reopened.GetSheet("Summary")).GetPivotTables());
    }

    /// <summary>
    /// 验证应用舍入后的金额及可精确表示的大整数保留数值缓存。
    /// </summary>
    [Fact]
    public void RoundedAmountsAndLargeIntegers_ShouldRemainNumericInCache()
    {
        var roundedAmount = (double)decimal.Round(123.456m, 2, MidpointRounding.AwayFromZero);
        var largeIntegers = new[] { 900719925474099d, 900719925474100d };
        var headers = new[] { "Region", "RoundedAmount", "LargeInteger" };
        using var workbook = CreateWorkbook(new[]
        {
            new object[] { "North", roundedAmount, largeIntegers[0] },
            new object[] { "South", -roundedAmount, largeIntegers[1] }
        }, headers);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 2
        }, new[] { Axis("Region") }, new[]
        {
            Value("RoundedAmount", "Amount"), Value("LargeInteger", "Large")
        });
        var source = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);
        var records = NpoiPivotTableWriter.Create(workbook, source, CancellationToken.None);
        using var output = new MemoryStream();

        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);

        var bytes = output.ToArray();
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XDocument ReadPart(string prefix)
        {
            var entry = Assert.Single(archive.Entries.Where(item =>
                item.FullName.Contains(prefix, StringComparison.Ordinal)
                && item.FullName.EndsWith(".xml", StringComparison.Ordinal)));
            using var stream = entry.Open();
            return XDocument.Load(stream);
        }

        var cacheFields = ReadPart("pivotCacheDefinition").Root
            .Element(spreadsheet + "cacheFields").Elements(spreadsheet + "cacheField").ToArray();
        var cacheRows = ReadPart("pivotCacheRecords").Root.Elements(spreadsheet + "r").ToArray();
        var amountField = Array.FindIndex(cacheFields,
            field => (string)field.Attribute("name") == "RoundedAmount");
        var largeField = Array.FindIndex(cacheFields,
            field => (string)field.Attribute("name") == "LargeInteger");

        Assert.Equal(new[] { roundedAmount, -roundedAmount }, cacheRows.Select(row =>
            ReadNumericCacheValue(row.Elements().ElementAt(amountField), amountField)));
        Assert.Equal(largeIntegers, cacheRows.Select(row =>
            ReadNumericCacheValue(row.Elements().ElementAt(largeField), largeField)));
        using var reopened = new XSSFWorkbook(new MemoryStream(bytes));
        Assert.Equal(roundedAmount, reopened.GetSheet("Data").GetRow(1).GetCell(1).NumericCellValue);
        Assert.Equal(largeIntegers[1], reopened.GetSheet("Data").GetRow(2).GetCell(2).NumericCellValue);

        double ReadNumericCacheValue(XElement recordCell, int fieldIndex)
        {
            var item = recordCell;
            if (recordCell.Name == spreadsheet + "x")
            {
                var field = cacheFields[fieldIndex];
                var itemIndex = int.Parse((string)recordCell.Attribute("v"), CultureInfo.InvariantCulture);
                item = field.Element(spreadsheet + "sharedItems").Elements().ElementAt(itemIndex);
            }

            Assert.Equal(spreadsheet + "n", item.Name);
            return double.Parse((string)item.Attribute("v"), CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// 验证日期单元格按 1900 与 1904 日期系统转换。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DateSerial_ShouldUseWorkbookDateSystem(bool date1904)
    {
        using var workbook = CreateWorkbook(new[]
        {
            new object[] { "North", 10d, "Retail", 1d, true, 1d, null, "x" }
        }, Headers, startRow: 3, startColumn: 2);
        if (date1904)
        {
            var properties = workbook.GetCTWorkbook().workbookPr ??= new CT_WorkbookPr();
            properties.date1904 = true;
            properties.date1904Specified = true;
        }
        var dateCell = workbook.GetSheet("Data").GetRow(4).GetCell(7);
        dateCell.CellStyle = CreateDateStyle(workbook);
        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = 4, EndColumn = 9
        }, rowFields: new[] { Axis("Region") }, valueFields: new[] { Value("Amount", "Total") });

        var plan = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);

        Assert.Equal(date1904, workbook.IsDate1904());
        Assert.Equal(DateUtil.GetJavaDate(1d, date1904), Assert.IsType<DateTime>(plan.Rows[0][5]));
    }

    /// <summary>
    /// 验证透视缓存导出与重开保留工作簿日期系统。
    /// </summary>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <param name="year">期望日期的年份。</param>
    /// <param name="month">期望日期的月份。</param>
    /// <param name="day">期望日期的日。</param>
    [Theory]
    [InlineData(false, 1900, 1, 1)]
    [InlineData(true, 1904, 1, 2)]
    public void DateCacheExport_ShouldPreserveWorkbookDateSystem(bool date1904, int year, int month, int day)
    {
        var expectedDate = new DateTime(year, month, day);
        using var workbook = CreateWorkbook(new[]
        {
            new object[] { "North", 10d, "Retail", 12d, false, 1d, null, "a" }
        }, Headers);
        if (date1904)
        {
            var properties = workbook.GetCTWorkbook().workbookPr ??= new CT_WorkbookPr();
            properties.date1904 = true;
            properties.date1904Specified = true;
        }
        workbook.GetSheet("Data").GetRow(1).GetCell(5).CellStyle = CreateDateStyle(workbook);

        var definition = AdvancedDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = Headers.Length - 1
        }, new[] { Axis("Region") }, new[]
        {
            new ExcelPivotValueDefinition
            {
                Field = "When", Caption = "Count dates", Aggregation = ExcelPivotAggregation.Count
            }
        });
        var source = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);
        Assert.Equal(expectedDate, Assert.IsType<DateTime>(source.Rows[0][5]));
        var records = NpoiPivotTableWriter.Create(workbook, source, CancellationToken.None);
        using var output = new MemoryStream();
        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);

        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var cacheEntry = Assert.Single(archive.Entries.Where(entry =>
            entry.FullName.Contains("pivotCacheDefinition", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)));
        using var cacheStream = cacheEntry.Open();
        var cache = XDocument.Load(cacheStream).Root;
        Assert.Equal("1", (string)cache.Attribute("recordCount"));
        var fields = cache.Element(spreadsheet + "cacheFields").Elements(spreadsheet + "cacheField").ToArray();
        var dateField = Assert.Single(fields.Where(field => (string)field.Attribute("name") == "When"));
        var dateItem = Assert.Single(dateField.Element(spreadsheet + "sharedItems").Elements());
        Assert.Equal("d", dateItem.Name.LocalName);
        Assert.Equal(expectedDate, DateTime.Parse((string)dateItem.Attribute("v"),
            System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind));

        var recordsEntry = Assert.Single(archive.Entries.Where(entry =>
            entry.FullName.Contains("pivotCacheRecords", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)));
        using var recordsStream = recordsEntry.Open();
        var cacheRecord = XDocument.Load(recordsStream).Root.Elements(spreadsheet + "r").Single();
        var dateReference = cacheRecord.Elements(spreadsheet + "x").ElementAt(5);
        Assert.Equal((uint)0, (uint)dateReference.Attribute("v"));

        using var reopened = new XSSFWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(date1904, reopened.IsDate1904());
        Assert.Equal(expectedDate, DateUtil.GetJavaDate(
            reopened.GetSheet("Data").GetRow(1).GetCell(5).NumericCellValue, reopened.IsDate1904()));
    }

    /// <summary>
    /// 验证展开预算的精确限额、多层轴、多个度量及物理边界。
    /// </summary>
    [Fact]
    public void LayoutBudget_ShouldAcceptExactAreaAndRejectOneCellLessOrPhysicalOverflow()
    {
        var rows = new[]
        {
            new object[] { "North", "A", "Jan", 10d },
            new object[] { "North", "B", "Feb", 20d },
            new object[] { "South", "A", "Jan", 30d }
        };
        using var workbook = CreateWorkbook(rows, new[] { "Region", "Category", "Month", "Amount" });
        var range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 3 };
        var multiAxis = AdvancedDefinition(range,
            rowFields: new[] { Axis("Region"), Axis("Category", ExcelPivotSubtotalMode.None) },
            columnFields: new[] { Axis("Month", ExcelPivotSubtotalMode.None) },
            valueFields: new[]
            {
                Value("Amount", "Total"),
                new ExcelPivotValueDefinition
                    { Field = "Amount", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
            }, maxExpandedCells: 64);

        var plan = NpoiPivotSourceValidator.Validate(workbook, multiAxis, CancellationToken.None);
        Assert.Equal(3, plan.Rows.Count);
        var oneCellShort = CloneDefinition(multiAxis, maxExpandedCells: 63);
        var budgetExceeded = CaptureLocation(workbook, oneCellShort);
        AssertLocation(budgetExceeded, 1, 1);
        Assert.Contains("64", budgetExceeded.Message);

        using var simpleWorkbook = CreateWorkbook(BasicRows, Headers);
        var exactRowEdge = LegacyDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 7
        }, targetRow: 1048572);
        Assert.Equal(3, NpoiPivotSourceValidator.Validate(simpleWorkbook, exactRowEdge, CancellationToken.None).Rows.Count);
        var rowOverflow = CloneDefinition(exactRowEdge, targetRow: 1048573);
        var rowOverflowError = CaptureLocation(simpleWorkbook, rowOverflow);
        AssertLocation(rowOverflowError, 1, 1);
        Assert.Contains("行列边界", rowOverflowError.Message);
        var exactColumnEdge = CloneDefinition(exactRowEdge, targetRow: 0, targetColumn: 16382);
        Assert.Equal(3, NpoiPivotSourceValidator.Validate(simpleWorkbook, exactColumnEdge,
            CancellationToken.None).Rows.Count);
        var columnOverflow = CloneDefinition(exactRowEdge, targetRow: 0, targetColumn: 16383);
        var columnOverflowError = CaptureLocation(simpleWorkbook, columnOverflow);
        AssertLocation(columnOverflowError, 1, 1);
        Assert.Contains("行列边界", columnOverflowError.Message);
    }

    /// <summary>
    /// 验证预取消、失败后重试及计划创建后的来源快照隔离。
    /// </summary>
    [Fact]
    public void Validation_ShouldHonorCancellationRetryAndSnapshotIsolation()
    {
        using var workbook = CreateWorkbook(BasicRows, Headers);
        var definition = LegacyDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 7
        });
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotSourceValidator.Validate(workbook, definition, canceled.Token));

        var data = workbook.GetSheet("Data");
        var original = data.GetRow(1).GetCell(7);
        original.SetCellFormula("1+2");
        var failure = CaptureLocation(workbook, definition);
        Assert.Equal(2, failure.RowIndex);
        Assert.Equal(8, failure.ColumnIndex);
        data.GetRow(1).RemoveCell(original);
        data.GetRow(1).CreateCell(7, CellType.String).SetCellValue("repaired");
        var plan = NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None);
        data.GetRow(1).GetCell(0).SetCellValue("Changed");

        Assert.Equal("North", Assert.IsType<string>(plan.Rows[0][0]));
        Assert.Equal("repaired", Assert.IsType<string>(plan.Rows[0][7]));
        Assert.Equal("Changed", data.GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 验证来源扫描进行中取消时立即停止读取。
    /// </summary>
    [Fact]
    public void Validation_ShouldHonorCancellationDuringSourceScan()
    {
        using var workbook = CreateWorkbook(BasicRows, Headers);
        var definition = LegacyDefinition(new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 7
        });
        using var canceled = new CancellationTokenSource();
        var readRows = new List<int>();

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotSourceValidator.Validate(workbook, definition, canceled.Token, row =>
            {
                readRows.Add(row);
                if (row == 2)
                    canceled.Cancel();
            }));

        Assert.Equal(canceled.Token, exception.CancellationToken);
        Assert.Equal(new[] { 1, 2 }, readRows);
        Assert.Equal("South", workbook.GetSheet("Data").GetRow(2).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 构造来源位于工作表偏移位置的真实 XLSX 工作簿。
    /// </summary>
    /// <param name="rows">按来源顺序排列的数据行。</param>
    /// <param name="headers">来源表头。</param>
    /// <param name="startRow">表头零基行号。</param>
    /// <param name="startColumn">表头零基列号。</param>
    /// <returns>包含 Data 工作表和真实类型单元格的工作簿。</returns>
    private static XSSFWorkbook CreateWorkbook(object[][] rows, string[] headers, int startRow = 0,
        int startColumn = 0)
    {
        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var dateStyle = CreateDateStyle(workbook);
        var header = sheet.CreateRow(startRow);
        for (var index = 0; index < headers.Length; index++)
            header.CreateCell(startColumn + index, CellType.String).SetCellValue(headers[index]);
        for (var rowIndex = 0; rowIndex < rows.Length; rowIndex++)
        {
            var row = sheet.CreateRow(startRow + rowIndex + 1);
            for (var column = 0; column < headers.Length; column++)
            {
                var cell = row.CreateCell(startColumn + column);
                switch (rows[rowIndex][column])
                {
                    case null: cell.SetCellType(CellType.Blank); break;
                    case string text: cell.SetCellValue(text); break;
                    case bool boolean: cell.SetCellValue(boolean); break;
                    case DateTime date: cell.SetCellValue(date); cell.CellStyle = dateStyle; break;
                    case double number: cell.SetCellValue(number); break;
                    case int number: cell.SetCellValue(number); break;
                    default: throw new ArgumentException("测试来源值类型未覆盖。");
                }
            }
        }
        return workbook;
    }

    /// <summary>
    /// 创建用于日期字面值识别的数值格式。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <returns>具有完整日期时间格式的单元格样式。</returns>
    private static ICellStyle CreateDateStyle(XSSFWorkbook workbook)
    {
        var style = workbook.CreateCellStyle();
        style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");
        return style;
    }

    /// <summary>
    /// 创建旧版单行 Sum 定义。
    /// </summary>
    /// <param name="range">含表头的来源区域。</param>
    /// <param name="rowField">行字段标题。</param>
    /// <param name="targetRow">透视目标零基行。</param>
    /// <param name="targetColumn">透视目标零基列。</param>
    /// <returns>供来源验证使用的透视定义。</returns>
    private static ExcelPivotTableDefinition LegacyDefinition(ExcelRangeDefinition range,
        string rowField = "Region", int targetRow = 0, int targetColumn = 0) => new()
    {
        Name = "SourceValidation",
        SourceSheet = "Data",
        SourceRange = range,
        TargetSheet = "Summary",
        TargetRow = targetRow,
        TargetColumn = targetColumn,
        RowField = rowField,
        ValueField = "Amount"
    };

    /// <summary>
    /// 创建多维轴、度量及筛选定义。
    /// </summary>
    /// <param name="range">含表头的来源区域。</param>
    /// <param name="rowFields">按层级顺序排列的行轴。</param>
    /// <param name="valueFields">按显示顺序排列的度量。</param>
    /// <param name="filters">报表筛选字段。</param>
    /// <param name="columnFields">按层级顺序排列的列轴。</param>
    /// <param name="maxExpandedCells">可选的展开预算。</param>
    /// <returns>供来源验证使用的高级透视定义。</returns>
    private static ExcelPivotTableDefinition AdvancedDefinition(ExcelRangeDefinition range,
        IReadOnlyList<ExcelPivotAxisFieldDefinition> rowFields,
        IReadOnlyList<ExcelPivotValueDefinition> valueFields,
        IReadOnlyList<ExcelPivotFilterDefinition> filters = null,
        IReadOnlyList<ExcelPivotAxisFieldDefinition> columnFields = null,
        long? maxExpandedCells = null) => new()
    {
        Name = "SourceValidation",
        SourceSheet = "Data",
        SourceRange = range,
        TargetSheet = "Summary",
        RowFields = rowFields,
        ColumnFields = columnFields ?? Array.Empty<ExcelPivotAxisFieldDefinition>(),
        ValueFields = valueFields,
        Filters = filters ?? Array.Empty<ExcelPivotFilterDefinition>(),
        MaxExpandedCells = maxExpandedCells
    };

    /// <summary>
    /// 创建行轴字段定义。
    /// </summary>
    /// <param name="field">来源表头。</param>
    /// <param name="subtotalMode">轴字段小计方式。</param>
    /// <returns>行轴字段。</returns>
    private static ExcelPivotAxisFieldDefinition Axis(string field,
        ExcelPivotSubtotalMode subtotalMode = ExcelPivotSubtotalMode.Automatic) => new()
    {
        Field = field,
        SubtotalMode = subtotalMode
    };

    /// <summary>
    /// 创建数值度量定义。
    /// </summary>
    /// <param name="field">来源表头。</param>
    /// <param name="caption">透视表显示标题。</param>
    /// <returns>使用 Sum 聚合的度量字段。</returns>
    private static ExcelPivotValueDefinition Value(string field, string caption) => new()
    {
        Field = field,
        Caption = caption,
        Aggregation = ExcelPivotAggregation.Sum
    };

    /// <summary>
    /// 创建单值报表筛选字段。
    /// </summary>
    /// <param name="field">来源表头。</param>
    /// <param name="value">筛选字面值。</param>
    /// <returns>只选择指定来源值的筛选字段。</returns>
    private static ExcelPivotFilterDefinition Filter(string field, object value) => new()
    {
        Field = field,
        Values = new[] { value }
    };

    /// <summary>
    /// 创建可调整预算或目标位置的定义副本。
    /// </summary>
    /// <param name="source">原始透视定义。</param>
    /// <param name="maxExpandedCells">替换后的单元格预算。</param>
    /// <param name="targetRow">替换后的目标零基行。</param>
    /// <param name="targetColumn">替换后的目标零基列。</param>
    /// <returns>配置其余内容与源定义相同的副本。</returns>
    private static ExcelPivotTableDefinition CloneDefinition(ExcelPivotTableDefinition source,
        long? maxExpandedCells = null, int? targetRow = null, int? targetColumn = null) => new()
    {
        Name = source.Name,
        SourceSheet = source.SourceSheet,
        SourceRange = source.SourceRange,
        TargetSheet = source.TargetSheet,
        TargetRow = targetRow ?? source.TargetRow,
        TargetColumn = targetColumn ?? source.TargetColumn,
        RowField = source.RowField,
        ValueField = source.ValueField,
        RowFields = source.RowFields,
        ColumnFields = source.ColumnFields,
        ValueFields = source.ValueFields,
        Filters = source.Filters,
        MaxExpandedCells = maxExpandedCells
    };

    /// <summary>
    /// 捕获含结构化原始坐标的来源配置异常。
    /// </summary>
    /// <param name="workbook">待验证工作簿。</param>
    /// <param name="definition">来源定义。</param>
    /// <returns>预检产生的配置异常。</returns>
    private static BingOfficesConfigurationException CaptureLocation(XSSFWorkbook workbook,
        ExcelPivotTableDefinition definition) => Assert.Throws<BingOfficesConfigurationException>(() =>
        NpoiPivotSourceValidator.Validate(workbook, definition, CancellationToken.None));

    /// <summary>
    /// 断言配置异常包含来源工作表、阶段和一基原始坐标。
    /// </summary>
    /// <param name="exception">来源预检异常。</param>
    /// <param name="row">期望的一基行号。</param>
    /// <param name="column">期望的一基列号。</param>
    private static void AssertLocation(BingOfficesConfigurationException exception, int row, int column)
    {
        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(row, exception.RowIndex);
        Assert.Equal(column, exception.ColumnIndex);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 断言来源异常精确指向给定的一基坐标。
    /// </summary>
    /// <param name="validator">来源验证入口。</param>
    /// <param name="workbook">待验证工作簿。</param>
    /// <param name="definition">来源定义。</param>
    /// <param name="row">期望的一基行号。</param>
    /// <param name="column">期望的一基列号。</param>
    /// <summary>
    /// 验证给定工作簿通过来源及容量预检。
    /// </summary>
    /// <param name="workbook">待验证工作簿。</param>
    /// <param name="definition">来源定义。</param>
    /// <returns>有效的来源计划。</returns>
    private static NpoiPivotSourcePlan Validate(XSSFWorkbook workbook,
        ExcelPivotTableDefinition definition) => NpoiPivotSourceValidator.Validate(workbook, definition,
        CancellationToken.None);

    /// <summary>
    /// 克隆来源行以构造隔离的无效值场景。
    /// </summary>
    /// <param name="rows">原始来源行。</param>
    /// <returns>逐行独立的可变值矩阵。</returns>
    private static object[][] CloneRows(object[][] rows) => rows.Select(row => (object[])row.Clone()).ToArray();
}
