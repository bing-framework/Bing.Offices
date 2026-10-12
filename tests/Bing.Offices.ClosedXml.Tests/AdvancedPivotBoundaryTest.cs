using System;
using System.IO;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证高级透视数据类型、过滤、预算和来源边界。
/// </summary>
public sealed class AdvancedPivotBoundaryTest
{
    /// <summary>
    /// SpreadsheetML 工作表主命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 验证数值聚合跳过空值，Count 保留混合来源类型。
    /// </summary>
    [Fact]
    public void Aggregations_ShouldPreserveNumericAndCountSourceTypes()
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Group";
            sheet.Cell("B1").Value = "Amount";
            sheet.Cell("C1").Value = "Mixed";
            sheet.Cell("D1").Value = "EmptyString";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = 10;
            sheet.Cell("C2").Value = "text";
            sheet.Cell("D2").Value = string.Empty;
            sheet.Cell("A3").Value = "North";
            sheet.Cell("C3").Value = 7;
            sheet.Cell("A4").Value = "South";
            sheet.Cell("B4").Value = 20;
            sheet.Cell("C4").Value = new DateTime(2025, 3, 4);
            sheet.Cell("A5").Value = "South";
            sheet.Cell("B5").Value = 30;
            sheet.Cell("C5").Value = true;
            sheet.Cell("A6").Value = "South";
        })) { Position = 0 };
        var request = CreateRequest(template, new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 5, EndColumn = 3
        }, new ExcelPivotAxisFieldDefinition { Field = "Group" }, new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Sum", Aggregation = ExcelPivotAggregation.Sum },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Average", Aggregation = ExcelPivotAggregation.Average },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Minimum", Aggregation = ExcelPivotAggregation.Minimum },
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Maximum", Aggregation = ExcelPivotAggregation.Maximum },
            new ExcelPivotValueDefinition { Field = "Mixed", Caption = "Count mixed", Aggregation = ExcelPivotAggregation.Count },
            new ExcelPivotValueDefinition { Field = "EmptyString", Caption = "Count empty string", Aggregation = ExcelPivotAggregation.Count }
        });

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
        Assert.Equal(new[] { XLPivotSummary.Sum, XLPivotSummary.Average, XLPivotSummary.Minimum,
            XLPivotSummary.Maximum, XLPivotSummary.Count, XLPivotSummary.Count },
            pivot.Values.Select(value => value.SummaryFormula));
        var records = ReadPart(output.ToArray(), "pivotCacheRecords");
        var mixedTypes = records.Root.Elements(SpreadsheetNs + "r")
            .Select(row => CacheItemType(output.ToArray(), row.Elements().ElementAt(2), 2))
            .ToArray();
        Assert.Equal(new[] { "s", "n", "d", "b", "m" }, mixedTypes);
        var emptyStringType = CacheItemType(output.ToArray(),
            records.Root.Elements(SpreadsheetNs + "r").First().Elements().ElementAt(3), 3);
        Assert.Equal("s", emptyStringType);
        var emptyStringTypes = records.Root.Elements(SpreadsheetNs + "r")
            .Select(row => CacheItemType(output.ToArray(), row.Elements().ElementAt(3), 3))
            .Distinct().ToArray();
        Assert.Equal(new[] { "s", "m" }, emptyStringTypes);

        var cacheFields = ReadPart(output.ToArray(), "pivotCacheDefinition").Root
            .Element(SpreadsheetNs + "cacheFields").Elements(SpreadsheetNs + "cacheField").ToArray();
        var mixedFlags = cacheFields[2].Element(SpreadsheetNs + "sharedItems");
        Assert.Equal("1", (string)mixedFlags.Attribute("containsString"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsSemiMixedTypes"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsNonDate"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsDate"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsBlank"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsMixedTypes"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsNumber"));
        Assert.Equal("0", (string)mixedFlags.Attribute("containsInteger"));
        Assert.Equal("0", (string)mixedFlags.Attribute("longText"));

        var amountFlags = cacheFields[1].Element(SpreadsheetNs + "sharedItems");
        Assert.Equal("0", (string)amountFlags.Attribute("containsString"));
        Assert.Equal("0", (string)amountFlags.Attribute("containsDate"));
        Assert.Equal("0", (string)amountFlags.Attribute("containsMixedTypes"));
        Assert.Equal("1", (string)amountFlags.Attribute("containsNumber"));
        Assert.Equal("1", (string)amountFlags.Attribute("containsInteger"));

        var emptyStringFlags = cacheFields[3].Element(SpreadsheetNs + "sharedItems");
        Assert.Equal("1", (string)emptyStringFlags.Attribute("containsString"));
        Assert.Equal("1", (string)emptyStringFlags.Attribute("containsBlank"));
        Assert.Equal("0", (string)emptyStringFlags.Attribute("containsMixedTypes"));
    }

    /// <summary>
    /// 验证应用舍入后的金额及可精确表示的大整数保留数值缓存。
    /// </summary>
    [Fact]
    public void RoundedAmountsAndLargeIntegers_ShouldRemainNumericInCache()
    {
        var roundedAmount = (double)decimal.Round(123.456m, 2, MidpointRounding.AwayFromZero);
        var largeIntegers = new[] { 900719925474099d, 900719925474100d };
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "RoundedAmount";
            sheet.Cell("C1").Value = "LargeInteger";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = roundedAmount;
            sheet.Cell("C2").Value = largeIntegers[0];
            sheet.Cell("A3").Value = "South";
            sheet.Cell("B3").Value = -roundedAmount;
            sheet.Cell("C3").Value = largeIntegers[1];
        })) { Position = 0 };
        var request = CreateRequest(template, CreateRange(0, 0, 2, 2),
            new ExcelPivotAxisFieldDefinition { Field = "Region" }, new[]
            {
                new ExcelPivotValueDefinition { Field = "RoundedAmount", Caption = "Amount" },
                new ExcelPivotValueDefinition { Field = "LargeInteger", Caption = "Large" }
            });
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, output);

        var bytes = output.ToArray();
        using var reopened = new XLWorkbook(new MemoryStream(bytes));
        Assert.Equal(roundedAmount, reopened.Worksheet("Data").Cell("B2").GetDouble());
        Assert.Equal(-roundedAmount, reopened.Worksheet("Data").Cell("B3").GetDouble());
        Assert.Equal(largeIntegers, new[]
        {
            reopened.Worksheet("Data").Cell("C2").GetDouble(),
            reopened.Worksheet("Data").Cell("C3").GetDouble()
        });

        var fields = ReadPart(bytes, "pivotCacheDefinition").Root
            .Element(SpreadsheetNs + "cacheFields").Elements(SpreadsheetNs + "cacheField").ToArray();
        var records = ReadPart(bytes, "pivotCacheRecords").Root.Elements(SpreadsheetNs + "r").ToArray();
        var amountField = Array.FindIndex(fields, field => (string)field.Attribute("name") == "RoundedAmount");
        var largeField = Array.FindIndex(fields, field => (string)field.Attribute("name") == "LargeInteger");

        Assert.Equal(new[] { roundedAmount, -roundedAmount }, records.Select(row =>
            ReadNumericCacheValue(bytes, row.Elements().ElementAt(amountField), amountField)));
        Assert.Equal(largeIntegers, records.Select(row =>
            ReadNumericCacheValue(bytes, row.Elements().ElementAt(largeField), largeField)));
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
    [InlineData(false, 2024, 6, 15)]
    [InlineData(true, 1904, 1, 2)]
    [InlineData(true, 2024, 6, 15)]
    public void DateCacheExport_ShouldPreserveWorkbookDateSystem(bool date1904, int year, int month, int day)
    {
        var expectedDate = new DateTime(year, month, day);
        var templateBytes = CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "When";
            sheet.Cell("C1").Value = "Amount";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = expectedDate;
            sheet.Cell("B2").Style.DateFormat.Format = "yyyy-mm-dd";
            sheet.Cell("C2").Value = 10;
        }, date1904);
        using var template = new MemoryStream(templateBytes) { Position = 0 };
        using (var sourceWorkbook = new XLWorkbook(new MemoryStream(template.ToArray())))
            Assert.Equal(expectedDate, sourceWorkbook.Worksheet("Data").Cell("B2").Value.GetDateTime());
        var request = CreateRequest(template, new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 2
        }, new ExcelPivotAxisFieldDefinition { Field = "Region" }, new[]
        {
            new ExcelPivotValueDefinition
            {
                Field = "When", Caption = "Count dates", Aggregation = ExcelPivotAggregation.Count
            }
        });

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        var cache = ReadPart(output.ToArray(), "pivotCacheDefinition").Root;
        var workbookDate1904 = (string)ReadPart(output.ToArray(), "workbook").Root
            .Element(SpreadsheetNs + "workbookPr").Attribute("date1904");
        Assert.Equal(date1904, workbookDate1904 == "1");
        var fields = cache.Element(SpreadsheetNs + "cacheFields").Elements(SpreadsheetNs + "cacheField").ToArray();
        var dateField = Assert.Single(fields.Where(field => (string)field.Attribute("name") == "When"));
        var dateItem = Assert.Single(dateField.Element(SpreadsheetNs + "sharedItems").Elements());
        Assert.Equal("d", dateItem.Name.LocalName);
        Assert.Equal(expectedDate, ParseCacheDate(dateItem));
        var records = ReadPart(output.ToArray(), "pivotCacheRecords").Root.Elements(SpreadsheetNs + "r").ToArray();
        var dateRecord = records.Single().Elements().ElementAt(1);
        if (dateRecord.Name == SpreadsheetNs + "x")
            Assert.Equal("0", (string)dateRecord.Attribute("v"));
        else
        {
            Assert.Equal(SpreadsheetNs + "d", dateRecord.Name);
            Assert.Equal(expectedDate, ParseCacheDate(dateRecord));
        }

        using var reopened = new XLWorkbook(new MemoryStream(output.ToArray()));
        Assert.Equal(expectedDate, reopened.Worksheet("Data").Cell("B2").Value.GetDateTime());
    }

    /// <summary>
    /// 验证混合类型共享项中的文本记录使用原始共享项索引。
    /// </summary>
    [Fact]
    public void PivotCacheNormalizer_ShouldPreserveMixedSharedItemIndexes()
    {
        using var source = new MemoryStream();
        using (var archive = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "xl/pivotCache/pivotCacheDefinition1.xml",
                "<pivotCacheDefinition xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
                + "<cacheFields><cacheField name='Mixed'><sharedItems><n v='7'/><s v='late'/></sharedItems>"
                + "</cacheField></cacheFields></pivotCacheDefinition>");
            WriteEntry(archive, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
                + "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                + "pivotCacheRecords' Target='pivotCacheRecords1.xml'/></Relationships>");
            WriteEntry(archive, "xl/pivotCache/pivotCacheRecords1.xml",
                "<pivotCacheRecords xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
                + "<r><n v='7'/></r><r><s v='late'/></r></pivotCacheRecords>");
        }
        source.Position = 0;
        using var destination = new MemoryStream();

        Bing.Offices.ClosedXml.Exports.PivotTables.ClosedXmlPivotCacheNormalizer.Normalize(
            source, destination, default);

        using var normalized = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        using var recordsStream = normalized.GetEntry("xl/pivotCache/pivotCacheRecords1.xml").Open();
        var records = XDocument.Load(recordsStream).Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal("n", records[0].Elements().Single().Name.LocalName);
        Assert.Equal("x", records[1].Elements().Single().Name.LocalName);
        Assert.Equal("1", (string)records[1].Elements().Single().Attribute("v"));
    }

    /// <summary>
    /// 验证未共享的缓存文本值仍参与字段类型标志计算。
    /// </summary>
    [Fact]
    public void PivotCacheNormalizer_ShouldInferSharedItemFlagsFromInlineRecords()
    {
        using var source = new MemoryStream();
        using (var archive = new ZipArchive(source, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "xl/pivotCache/pivotCacheDefinition1.xml",
                "<pivotCacheDefinition xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
                + "<cacheFields><cacheField name='Mixed'><sharedItems><n v='7'/></sharedItems></cacheField>"
                + "<cacheField name='Item'><sharedItems/></cacheField></cacheFields></pivotCacheDefinition>");
            WriteEntry(archive, "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
                "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>"
                + "<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                + "pivotCacheRecords' Target='pivotCacheRecords1.xml'/></Relationships>");
            WriteEntry(archive, "xl/pivotCache/pivotCacheRecords1.xml",
                "<pivotCacheRecords xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>"
                + "<r><n v='7'/><s v='A'/></r><r><s v='late'/><s v='B'/></r></pivotCacheRecords>");
        }
        source.Position = 0;
        using var destination = new MemoryStream();

        Bing.Offices.ClosedXml.Exports.PivotTables.ClosedXmlPivotCacheNormalizer.Normalize(
            source, destination, default);

        using var normalized = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        using var definitionStream = normalized.GetEntry("xl/pivotCache/pivotCacheDefinition1.xml").Open();
        var fields = XDocument.Load(definitionStream).Root.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        var mixedFlags = fields[0].Element(SpreadsheetNs + "sharedItems");
        Assert.Equal("1", (string)mixedFlags.Attribute("containsString"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsNumber"));
        Assert.Equal("1", (string)mixedFlags.Attribute("containsMixedTypes"));
        Assert.Equal("0", (string)mixedFlags.Attribute("containsInteger"));

        var itemFlags = fields[1].Element(SpreadsheetNs + "sharedItems");
        Assert.Equal("1", (string)itemFlags.Attribute("containsString"));
        Assert.Equal("1", (string)itemFlags.Attribute("containsSemiMixedTypes"));
        Assert.Equal("1", (string)itemFlags.Attribute("containsNonDate"));
        Assert.Equal("0", (string)itemFlags.Attribute("containsNumber"));
        Assert.Equal("0", (string)itemFlags.Attribute("containsMixedTypes"));
    }

    /// <summary>
    /// 验证数值聚合拒绝文本、日期、布尔及全空来源。
    /// </summary>
    /// <param name="valueKind">非法或全空数值列的内容类型。</param>
    [Theory]
    [InlineData("text")]
    [InlineData("date")]
    [InlineData("bool")]
    [InlineData("empty")]
    public void NumericAggregations_ShouldRejectNonNumericAndAllEmptyValues(string valueKind)
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("C4").Value = "Region";
            sheet.Cell("D4").Value = "Amount";
            sheet.Cell("C5").Value = "North";
            switch (valueKind)
            {
                case "text": sheet.Cell("D5").Value = "12"; break;
                case "date": sheet.Cell("D5").Value = new DateTime(2025, 1, 2); break;
                case "bool": sheet.Cell("D5").Value = true; break;
                case "empty": sheet.Cell("D5").Clear(); break;
            }
            if (valueKind != "empty")
            {
                sheet.Cell("C6").Value = "South";
                sheet.Cell("D6").Value = 5;
            }
        })) { Position = 0 };
        var request = CreateRequest(template, new ExcelRangeDefinition
        {
            StartRow = 3, StartColumn = 2, EndRow = valueKind == "empty" ? 4 : 5, EndColumn = 3
        }, new ExcelPivotAxisFieldDefinition { Field = "Region" }, new[]
        {
            new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" }
        });
        using var output = SeededStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(5, exception.RowIndex);
        Assert.Equal(4, exception.ColumnIndex);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(new byte[] { 7, 11, 19, 23 }, output.ToArray());
    }

    /// <summary>
    /// 验证数值、日期、布尔和空值筛选写入完整报表筛选与缓存。
    /// </summary>
    [Fact]
    public void TypedFilters_ShouldSerializeAllSelectionsAndKeepSourceRows()
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Amount";
            sheet.Cell("C1").Value = "Numeric";
            sheet.Cell("D1").Value = "When";
            sheet.Cell("E1").Value = "Active";
            sheet.Cell("F1").Value = "Optional";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = 10;
            sheet.Cell("C2").Value = 3;
            sheet.Cell("D2").Value = new DateTime(2025, 1, 1);
            sheet.Cell("E2").Value = true;
            sheet.Cell("A3").Value = "South";
            sheet.Cell("B3").Value = 20;
            sheet.Cell("C3").Value = 8;
            sheet.Cell("D3").Value = new DateTime(2025, 1, 2);
            sheet.Cell("E3").Value = false;
            sheet.Cell("F3").Value = "filled";
        })) { Position = 0 };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "TypedFilterPivot", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 5 },
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            Filters = new[]
            {
                new ExcelPivotFilterDefinition { Field = "Numeric", Values = new object[] { 3 } },
                new ExcelPivotFilterDefinition { Field = "When", Values = new object[] { new DateTime(2025, 1, 2) } },
                new ExcelPivotFilterDefinition { Field = "Active", Values = new object[] { false } },
                new ExcelPivotFilterDefinition { Field = "Optional", Values = new object[] { null } }
            }
        };
        var request = CreateRequest(template, definition);

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
        Assert.Equal(new[] { "Numeric", "When", "Active", "Optional" },
            pivot.ReportFilters.Select(filter => filter.SourceName));
        var archive = ReadPart(output.ToArray(), "pivotCacheRecords");
        Assert.Equal(2, archive.Root.Elements(SpreadsheetNs + "r").Count());
        var fieldTypes = Enumerable.Range(2, 4).Select(index => archive.Root.Elements(SpreadsheetNs + "r")
            .Select(row => CacheItemType(output.ToArray(), row.Elements().ElementAt(index), index)).Distinct().ToArray())
            .ToArray();
        Assert.Contains("n", fieldTypes[0]);
        Assert.Contains("d", fieldTypes[1]);
        Assert.Contains("b", fieldTypes[2]);
        Assert.Contains("m", fieldTypes[3]);
        var tableXml = ReadPart(output.ToArray(), "pivotTable");
        var cacheXml = ReadPart(output.ToArray(), "pivotCacheDefinition");
        var cacheFields = cacheXml.Root.Element(SpreadsheetNs + "cacheFields").Elements().ToArray();
        var pivotFields = tableXml.Root.Element(SpreadsheetNs + "pivotFields").Elements().ToArray();
        var pages = tableXml.Root.Element(SpreadsheetNs + "pageFields").Elements().ToArray();
        Assert.Equal(new[] { "2", "3", "4", "5" }, pages.Select(page => (string)page.Attribute("fld")));
        var selections = pages.Select(page =>
        {
            var fieldIndex = int.Parse((string)page.Attribute("fld"));
            var itemIndex = int.Parse((string)page.Attribute("item"));
            var cacheIndex = int.Parse((string)pivotFields[fieldIndex].Element(SpreadsheetNs + "items")
                .Elements().ElementAt(itemIndex).Attribute("x"));
            return cacheFields[fieldIndex].Element(SpreadsheetNs + "sharedItems").Elements().ElementAt(cacheIndex);
        }).ToArray();
        Assert.Equal(new[] { "n", "d", "b", "m" }, selections.Select(item => item.Name.LocalName));
        Assert.Equal("3", (string)selections[0].Attribute("v"));
        Assert.Equal(new DateTime(2025, 1, 2), DateTime.Parse((string)selections[1].Attribute("v"),
            System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("0", (string)selections[2].Attribute("v"));
        Assert.Null(selections[3].Attribute("v"));
    }

    /// <summary>
    /// 验证筛选类型不匹配时报告来源表头的原始坐标。
    /// </summary>
    [Fact]
    public void FilterTypeMismatch_ShouldReportOriginalHeaderCoordinate()
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("C4").Value = "Region";
            sheet.Cell("D4").Value = "Amount";
            sheet.Cell("E4").Value = "StatusCode";
            sheet.Cell("C5").Value = "North";
            sheet.Cell("D5").Value = 4;
            sheet.Cell("E5").Value = 12;
        })) { Position = 0 };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "Mismatch", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 3, StartColumn = 2, EndRow = 4, EndColumn = 4 },
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            Filters = new[] { new ExcelPivotFilterDefinition { Field = "StatusCode", Values = new object[] { "12" } } }
        };
        using var output = SeededStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(CreateRequest(template, definition), output));

        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(4, exception.RowIndex);
        Assert.Equal(5, exception.ColumnIndex);
        Assert.Equal(new byte[] { 7, 11, 19, 23 }, output.ToArray());
    }

    /// <summary>
    /// 验证仅列轴、多层列轴和多个报表筛选均完整写入。
    /// </summary>
    [Fact]
    public void ColumnOnlyAxisAndMultipleFilters_ShouldPreserveFieldOrder()
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Category";
            sheet.Cell("C1").Value = "Month";
            sheet.Cell("D1").Value = "Status";
            sheet.Cell("E1").Value = "Amount";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = "Food";
            sheet.Cell("C2").Value = "Jan";
            sheet.Cell("D2").Value = "Open";
            sheet.Cell("E2").Value = 10;
            sheet.Cell("A3").Value = "South";
            sheet.Cell("B3").Value = "Tools";
            sheet.Cell("C3").Value = "Feb";
            sheet.Cell("D3").Value = "Closed";
            sheet.Cell("E3").Value = 20;
        })) { Position = 0 };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "ColumnOnly", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 4 },
            ColumnFields = new[]
            {
                new ExcelPivotAxisFieldDefinition { Field = "Region" },
                new ExcelPivotAxisFieldDefinition { Field = "Category", SubtotalMode = ExcelPivotSubtotalMode.None }
            },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            Filters = new[]
            {
                new ExcelPivotFilterDefinition { Field = "Month", Values = new object[] { "Jan" } },
                new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { "Open" } }
            }
        };

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(CreateRequest(template, definition), output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray()));
        var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
        Assert.Empty(pivot.RowLabels);
        Assert.Equal(new[] { "Region", "Category" }, pivot.ColumnLabels.Select(field => field.SourceName));
        Assert.Equal(new[] { "Month", "Status" }, pivot.ReportFilters.Select(filter => filter.SourceName));
        var pivotXml = ReadPart(output.ToArray(), "pivotTable");
        Assert.Equal(2, pivotXml.Root.Element(SpreadsheetNs + "colFields").Elements().Count());
        Assert.Equal(2, pivotXml.Root.Element(SpreadsheetNs + "pageFields").Elements().Count());
    }

    /// <summary>
    /// 验证高级与旧版展开预算的边界，以及目标行列格式上限。
    /// </summary>
    [Fact]
    public void ExpandedBudgetAndPhysicalBounds_ShouldAcceptExactAndRejectOverflow()
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("A1").Value = "Region";
            sheet.Cell("B1").Value = "Amount";
            sheet.Cell("A2").Value = "North";
            sheet.Cell("B2").Value = 1;
            sheet.Cell("A3").Value = "South";
            sheet.Cell("B3").Value = 2;
        })) { Position = 0 };

        using var exact = new MemoryStream();
        new ClosedXmlExcelExporter().Export(CreateRequest(template, CreateRange(0, 0, 2, 1),
            new ExcelPivotAxisFieldDefinition { Field = "Region" },
            new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } }, budget: 8), exact);
        AssertCompleteBudgetWorkbook(exact.ToArray());

        using var physicalExact = new MemoryStream();
        new ClosedXmlExcelExporter().Export(CreateRequest(template, CreateRange(0, 0, 2, 1),
            new ExcelPivotAxisFieldDefinition { Field = "Region" },
            new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            targetRow: 1048572, targetColumn: 16382), physicalExact);
        AssertCompleteBudgetWorkbook(physicalExact.ToArray());

        AssertPreflightFailure(CreateRequest(template, CreateRange(0, 0, 2, 1),
            new ExcelPivotAxisFieldDefinition { Field = "Region" },
            new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } }, budget: 7), 1, 1);
        AssertPreflightFailure(CreateRequest(template, CreateRange(0, 0, 2, 1),
            new ExcelPivotAxisFieldDefinition { Field = "Region" },
            new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            targetRow: 1048573), 1, 1);
        AssertPreflightFailure(CreateRequest(template, new ExcelPivotTableDefinition
        {
            Name = "ColumnOverflow", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = CreateRange(0, 0, 2, 1), TargetColumn = 16382,
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[]
            {
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" },
                new ExcelPivotValueDefinition { Field = "Amount", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
            }
        }), 1, 1);

        var legacyExact = CreateLegacyRequest(template, budget: 8);
        using var legacyOutput = new MemoryStream();
        new ClosedXmlExcelExporter().Export(legacyExact, legacyOutput);
        AssertCompleteBudgetWorkbook(legacyOutput.ToArray());
        AssertPreflightFailure(CreateLegacyRequest(template, budget: 7), 1, 1);
    }

    /// <summary>
    /// 验证高级来源区域内非度量公式和 Excel 错误都按原坐标拒绝。
    /// </summary>
    /// <param name="errorValue">是否写入 Excel 错误而非公式。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormulaOrErrorInNonMetricSourceField_ShouldRejectAtOriginalCoordinate(bool errorValue)
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("C4").Value = "Region";
            sheet.Cell("D4").Value = "Amount";
            sheet.Cell("E4").Value = "Memo";
            sheet.Cell("C5").Value = "North";
            sheet.Cell("D5").Value = 10;
            if (errorValue)
                sheet.Cell("E5").Value = XLError.DivisionByZero;
            else
                sheet.Cell("E5").FormulaA1 = "1+1";
        })) { Position = 0 };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "NonMetricSource", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 3, StartColumn = 2, EndRow = 4, EndColumn = 4 },
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } }
        };

        AssertPreflightFailure(CreateRequest(template, definition), 5, 5);
    }

    /// <summary>
    /// 创建带测试数据的 XLSX 模板。
    /// </summary>
    /// <param name="configure">配置来源工作表内容的操作。</param>
    /// <param name="date1904">是否使用 1904 日期系统。</param>
    /// <returns>模板文件字节。</returns>
    private static byte[] CreateTemplate(Action<IXLWorksheet> configure, bool date1904 = false)
    {
        using var workbook = new XLWorkbook { Use1904DateSystem = date1904 };
        configure(workbook.Worksheets.Add("Data"));
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建高级透视请求。
    /// </summary>
    /// <param name="template">模板工作簿。</param>
    /// <param name="range">来源区域。</param>
    /// <param name="rowField">行轴字段；可为空以创建仅列轴透视。</param>
    /// <param name="values">度量定义。</param>
    /// <param name="budget">可选保守展开上限。</param>
    /// <param name="targetRow">零基目标行。</param>
    /// <param name="targetColumn">零基目标列。</param>
    /// <returns>已验证的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template, ExcelRangeDefinition range,
        ExcelPivotAxisFieldDefinition rowField, ExcelPivotValueDefinition[] values,
        long? budget = null, int targetRow = 0, int targetColumn = 0) => CreateRequest(template,
            new ExcelPivotTableDefinition
            {
                Name = "BoundaryPivot", SourceSheet = "Data", TargetSheet = "Summary",
                SourceRange = range, TargetRow = targetRow, TargetColumn = targetColumn,
                RowFields = rowField == null ? Array.Empty<ExcelPivotAxisFieldDefinition>() : new[] { rowField },
                ValueFields = values,
                MaxExpandedCells = budget
            });

    /// <summary>
    /// 创建已配置的透视请求。
    /// </summary>
    /// <param name="template">模板工作簿。</param>
    /// <param name="definition">透视表定义。</param>
    /// <returns>已验证的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template,
        ExcelPivotTableDefinition definition) => ExcelExport.Workbook(builder => builder
        .UseTemplate(Reset(template), leaveOpen: true)
        .AddSheet("Data", Array.Empty<AdvancedPivotBoundaryRow>(), sheet => sheet.HeaderRowIndex(10))
        .PivotTable(definition));

    /// <summary>
    /// 创建旧版单行求和透视请求。
    /// </summary>
    /// <param name="template">模板工作簿。</param>
    /// <param name="budget">保守展开上限。</param>
    /// <returns>旧版字段布局请求。</returns>
    private static ExcelWorkbookExportRequest CreateLegacyRequest(Stream template, long budget) =>
        ExcelExport.Workbook(builder => builder.UseTemplate(Reset(template), leaveOpen: true)
            .AddSheet("Data", Array.Empty<AdvancedPivotBoundaryRow>(), sheet => sheet.HeaderRowIndex(10))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "LegacyBoundary", SourceSheet = "Data", TargetSheet = "Summary",
                SourceRange = CreateRange(0, 0, 2, 1), RowField = "Region", ValueField = "Amount",
                MaxExpandedCells = budget
            }));

    /// <summary>
    /// 创建来源区域定义。
    /// </summary>
    /// <param name="startRow">零基起始行。</param>
    /// <param name="startColumn">零基起始列。</param>
    /// <param name="endRow">零基结束行。</param>
    /// <param name="endColumn">零基结束列。</param>
    /// <returns>指定的来源区域。</returns>
    private static ExcelRangeDefinition CreateRange(int startRow, int startColumn, int endRow, int endColumn) => new()
    {
        StartRow = startRow, StartColumn = startColumn, EndRow = endRow, EndColumn = endColumn
    };

    /// <summary>
    /// 验证预算边界工作簿的完整来源与缓存。
    /// </summary>
    /// <param name="bytes">成功序列化的工作簿。</param>
    private static void AssertCompleteBudgetWorkbook(byte[] bytes)
    {
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        Assert.Equal(new[] { "Data", "Summary" }, workbook.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "Region", "Amount", "North", "1", "South", "2" },
            workbook.Worksheet("Data").Range("A1:B3").Cells().Select(cell => cell.GetString()));
        Assert.Single(workbook.Worksheet("Summary").PivotTables);
        var records = ReadPart(bytes, "pivotCacheRecords").Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(2, records.Length);
        Assert.All(records, row => Assert.Equal(2, row.Elements().Count()));
        var fields = ReadPart(bytes, "pivotCacheDefinition").Root.Element(SpreadsheetNs + "cacheFields")
            .Elements().ToArray();
        Assert.Equal(new[] { "Region", "Amount" }, fields.Select(field => (string)field.Attribute("name")));
        Assert.Equal(new[] { "North", "South" }, records.Select(row => row.Elements().First())
            .Select(item => item.Name.LocalName == "x"
                ? (string)fields[0].Element(SpreadsheetNs + "sharedItems").Elements()
                    .ElementAt(int.Parse((string)item.Attribute("v"))).Attribute("v")
                : (string)item.Attribute("v")));
        Assert.Equal(new[] { "1", "2" }, records.Select(row => (string)row.Elements().Last().Attribute("v")));
    }

    /// <summary>
    /// 验证来源预检失败的坐标和输出流保护。
    /// </summary>
    /// <param name="request">包含非法来源或边界的请求。</param>
    /// <param name="row">预期一基行号。</param>
    /// <param name="column">预期一基列号。</param>
    private static void AssertPreflightFailure(ExcelWorkbookExportRequest request, int row, int column)
    {
        using var output = SeededStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(row, exception.RowIndex);
        Assert.Equal(column, exception.ColumnIndex);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(new byte[] { 7, 11, 19, 23 }, output.ToArray());
    }

    /// <summary>
    /// 创建带哨兵字节的目标流。
    /// </summary>
    /// <returns>位于起点且包含固定哨兵内容的流。</returns>
    private static MemoryStream SeededStream()
    {
        var output = new MemoryStream();
        output.Write(new byte[] { 7, 11, 19, 23 });
        output.Position = 0;
        return output;
    }

    /// <summary>
    /// 读取透视相关的唯一 OOXML 部件。
    /// </summary>
    /// <param name="bytes">完整工作簿字节。</param>
    /// <param name="prefix">部件文件名前缀。</param>
    /// <returns>已解析的部件文档。</returns>
    private static XDocument ReadPart(byte[] bytes, string prefix)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes));
        var entry = Assert.Single(archive.Entries.Where(item =>
            item.FullName.Split('/').Last().StartsWith(prefix, StringComparison.Ordinal)
            && item.FullName.EndsWith(".xml", StringComparison.Ordinal)));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 读取透视缓存中的日期值。
    /// </summary>
    /// <param name="item">缓存日期项。</param>
    /// <returns>解析后的日期。</returns>
    private static DateTime ParseCacheDate(XElement item) => DateTime.Parse(
        (string)item.Attribute("v") ?? item.Value,
        System.Globalization.CultureInfo.InvariantCulture,
        System.Globalization.DateTimeStyles.RoundtripKind);

    /// <summary>
    /// 将文本写入 ZIP 部件。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="name">部件路径。</param>
    /// <param name="content">部件内容。</param>
    private static void WriteEntry(ZipArchive archive, string name, string content)
    {
        var entry = archive.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    /// <summary>
    /// 获取缓存记录中某字段对应的缓存项类型。
    /// </summary>
    /// <param name="workbookBytes">完整工作簿字节。</param>
    /// <param name="recordCell">记录中的字段引用。</param>
    /// <param name="fieldIndex">来源字段零基索引。</param>
    /// <returns>SpreadsheetML 缓存项的元素名。</returns>
    private static string CacheItemType(byte[] workbookBytes, XElement recordCell, int fieldIndex)
    {
        if (recordCell.Name != SpreadsheetNs + "x")
            return recordCell.Name.LocalName;
        var cache = ReadPart(workbookBytes, "pivotCacheDefinition");
        var cacheField = cache.Root.Element(SpreadsheetNs + "cacheFields").Elements().ElementAt(fieldIndex);
        var itemIndex = int.Parse((string)recordCell.Attribute("v"));
        return cacheField.Element(SpreadsheetNs + "sharedItems").Elements().ElementAt(itemIndex)
            .Name.LocalName;
    }

    /// <summary>
    /// 读取缓存记录中的数值字面值。
    /// </summary>
    /// <param name="workbookBytes">完整工作簿字节。</param>
    /// <param name="recordCell">记录中的字段值或共享项引用。</param>
    /// <param name="fieldIndex">来源字段零基索引。</param>
    /// <returns>缓存中的数值。</returns>
    private static double ReadNumericCacheValue(byte[] workbookBytes, XElement recordCell, int fieldIndex)
    {
        var item = recordCell;
        if (recordCell.Name == SpreadsheetNs + "x")
        {
            var cache = ReadPart(workbookBytes, "pivotCacheDefinition");
            var field = cache.Root.Element(SpreadsheetNs + "cacheFields").Elements().ElementAt(fieldIndex);
            var itemIndex = int.Parse((string)recordCell.Attribute("v"), CultureInfo.InvariantCulture);
            item = field.Element(SpreadsheetNs + "sharedItems").Elements().ElementAt(itemIndex);
        }

        Assert.Equal(SpreadsheetNs + "n", item.Name);
        return double.Parse((string)item.Attribute("v"), CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 将模板流定位到起点后返回同一流。
    /// </summary>
    /// <param name="template">模板工作簿流。</param>
    /// <returns>定位到起点的模板流。</returns>
    private static Stream Reset(Stream template)
    {
        template.Position = 0;
        return template;
    }

}
