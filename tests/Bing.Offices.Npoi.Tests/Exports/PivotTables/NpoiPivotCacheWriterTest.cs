using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
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
/// 验证 NPOI 透视缓存字段、记录写入和输入边界。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotCacheWriterTest
{
    /// <summary>
    /// 六个缓存字段的固定顺序。
    /// </summary>
    private static readonly string[] Headers = { "Warehouse", "Amount", "Item", "Enabled", "Date", "Optional" };

    /// <summary>
    /// 覆盖文本、数值、布尔、日期、空字符串和真空白的完整来源矩阵。
    /// </summary>
    private static readonly object[][] SourceRows =
    {
        new object[] { "North", 10d, "Widget", true, new DateTime(2026, 1, 2, 13, 45, 0), "memo" },
        new object[] { "South", 20d, string.Empty, false, new DateTime(2026, 1, 3, 9, 15, 0), null },
        new object[] { "North", 30d, "Screw", true, new DateTime(2026, 1, 2, 13, 45, 0), null }
    };

    /// <summary>
    /// 验证来源矩阵形成类型化共享项、去重索引和完整记录。
    /// </summary>
    [Fact]
    public void Complete_ShouldBuildTypedSharedItemsAndRecordIndexes()
    {
        using var workbook = CreateWorkbook(out var pivot);
        var records = NpoiPivotCacheWriter.Complete(pivot, Headers, ReadOnlyRows(SourceRows),
            CancellationToken.None);
        var fields = pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition().cacheFields.cacheField;

        Assert.Equal("xl/pivotCache/pivotCacheRecords1.xml", records.PartName);
        Assert.Equal(Headers.Length, records.FieldCount);
        Assert.Equal(new[] { 2, 3, 3, 2, 2, 2 }, records.SharedItemCounts);
        Assert.Equal(new uint[][]
        {
            new uint[] { 0, 0, 0, 0, 0, 0 },
            new uint[] { 1, 1, 1, 1, 1, 1 },
            new uint[] { 0, 2, 2, 0, 0, 1 }
        }, records.Rows.Select(row => row.ToArray()).ToArray());
        Assert.Equal(Headers.Length, fields.Count);
        Assert.Equal(Headers, fields.Select(field => field.name));
        Assert.Equal((uint)SourceRows.Length, pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition().recordCount);
        Assert.Collection(fields[0].sharedItems.Items,
            item => Assert.IsType<CT_String>(item), item => Assert.IsType<CT_String>(item));
        Assert.Collection(fields[1].sharedItems.Items,
            item => Assert.IsType<CT_Number>(item), item => Assert.IsType<CT_Number>(item),
            item => Assert.IsType<CT_Number>(item));
        Assert.Collection(fields[2].sharedItems.Items,
            item => Assert.IsType<CT_String>(item), item => Assert.IsType<CT_String>(item),
            item => Assert.IsType<CT_String>(item));
        Assert.Equal(string.Empty, Assert.IsType<CT_String>(fields[2].sharedItems.Items[1]).v);
        Assert.Collection(fields[3].sharedItems.Items,
            item => Assert.IsType<CT_Boolean>(item), item => Assert.IsType<CT_Boolean>(item));
        Assert.Collection(fields[4].sharedItems.Items,
            item => Assert.IsType<CT_DateTime>(item), item => Assert.IsType<CT_DateTime>(item));
        Assert.Collection(fields[5].sharedItems.Items,
            item => Assert.IsType<CT_String>(item), item => Assert.IsType<CT_Missing>(item));
        Assert.True(fields[5].sharedItems.containsBlank);
        Assert.True(fields[2].sharedItems.containsString);
        Assert.True(fields[1].sharedItems.containsNumber);
        Assert.False(fields[3].sharedItems.containsMixedTypes);
        Assert.True(fields[4].sharedItems.containsDate);
        var rowFieldItems = pivot.GetCTPivotTableDefinition().pivotFields.pivotField[0].items.item;
        Assert.Equal(3, rowFieldItems.Count);
        Assert.Equal("data", rowFieldItems[0].t.ToString());
        Assert.Equal((uint)0, rowFieldItems[0].x);
        Assert.Equal("data", rowFieldItems[1].t.ToString());
        Assert.Equal((uint)1, rowFieldItems[1].x);
        Assert.Equal("default", rowFieldItems[2].t.ToString());
    }

    /// <summary>
    /// 验证混合类型共享项、空白索引及长文本标志写入完整 XLSX XML。
    /// </summary>
    [Fact]
    public void Complete_ShouldSerializeMixedSharedItemTypesAndFlags()
    {
        var longText = "\"中文\"\n" + new string('界', 300);
        var date = new DateTime(2026, 4, 5, 13, 45, 30);
        object[] values = { 42d, longText, true, date, null, 42d, null, string.Empty };
        using var workbook = CreateSingleFieldWorkbook("Mixed", values, out var pivot);
        var records = NpoiPivotCacheWriter.Complete(pivot, new[] { "Mixed" },
            ReadOnlyRows(values.Select(value => new[] { value }).ToArray()), CancellationToken.None);
        using var output = new MemoryStream();
        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var mainNs = (XNamespace)"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var packageRelationshipsNs = (XNamespace)"http://schemas.openxmlformats.org/package/2006/relationships";
        var workbookRelationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var definitionRelationship = Assert.Single(workbookRelationships.Root
            .Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true));
        var definitionPart = ResolvePart("xl/workbook.xml", (string)definitionRelationship.Attribute("Target"));
        var definition = ReadXml(archive, definitionPart);
        var cacheField = Assert.Single(definition.Root.Element(mainNs + "cacheFields")
            .Elements(mainNs + "cacheField"));
        var sharedItems = cacheField.Element(mainNs + "sharedItems");
        var shared = sharedItems.Elements().ToArray();
        Assert.Equal("Mixed", (string)cacheField.Attribute("name"));
        Assert.Equal(new[] { "n", "s", "b", "d", "m", "s" }, shared.Select(item => item.Name.LocalName));
        Assert.Equal(42d, double.Parse((string)shared[0].Attribute("v"), CultureInfo.InvariantCulture));
        Assert.Equal(longText, DecodeXstring((string)shared[1].Attribute("v")));
        Assert.Equal("1", (string)shared[2].Attribute("v"));
        Assert.Equal(date, DateTime.Parse((string)shared[3].Attribute("v"), CultureInfo.InvariantCulture,
            DateTimeStyles.RoundtripKind));
        Assert.Null(shared[4].Attribute("v"));
        Assert.Equal(string.Empty, DecodeXstring((string)shared[5].Attribute("v")));
        Assert.True(XmlFlag(sharedItems, "containsMixedTypes"));
        Assert.True(XmlFlag(sharedItems, "containsSemiMixedTypes"));
        Assert.True(XmlFlag(sharedItems, "containsString"));
        Assert.True(XmlFlag(sharedItems, "containsNumber"));
        Assert.True(XmlFlag(sharedItems, "containsDate"));
        Assert.True(XmlFlag(sharedItems, "containsNonDate"));
        Assert.True(XmlFlag(sharedItems, "containsBlank"));
        Assert.True(XmlFlag(sharedItems, "longText"));
        Assert.False(XmlFlag(sharedItems, "containsInteger"));

        var cacheRecords = ReadXml(archive, records.PartName).Root.Elements(mainNs + "r").ToArray();
        Assert.Equal(new[] { 0U, 1U, 2U, 3U, 4U, 0U, 4U, 5U }, cacheRecords
            .Select(record => uint.Parse((string)Assert.Single(record.Elements(mainNs + "x")).Attribute("v"),
                CultureInfo.InvariantCulture)));
    }

    /// <summary>
    /// 验证数值缓存字段显式写出与实际共享项一致的类型标志。
    /// </summary>
    [Fact]
    public void ContentWriter_ShouldWriteExplicitNumericSharedItemFlags()
    {
        using var workbook = CreateSingleFieldWorkbook("Amount", new object[] { 1d, 2d, 3d }, out var pivot);
        var records = NpoiPivotCacheWriter.Complete(pivot, new[] { "Amount" },
            ReadOnlyRows(new[]
            {
                new object[] { 1d },
                new object[] { 2d },
                new object[] { 3d }
            }), CancellationToken.None);
        using var output = new MemoryStream();

        NpoiSheetContentWriter.Write(workbook,
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>(),
            new[] { records }, output, CancellationToken.None);

        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var mainNs = (XNamespace)"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var packageRelationshipsNs = (XNamespace)"http://schemas.openxmlformats.org/package/2006/relationships";
        var workbookRelationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var definitionRelationship = Assert.Single(workbookRelationships.Root
            .Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true));
        var definitionPart = ResolvePart("xl/workbook.xml", (string)definitionRelationship.Attribute("Target"));
        var definition = ReadXml(archive, definitionPart);
        var sharedItems = Assert.Single(definition.Root.Element(mainNs + "cacheFields")
            .Elements(mainNs + "cacheField")).Element(mainNs + "sharedItems");

        Assert.Equal("1", (string)sharedItems.Attribute("containsNumber"));
        Assert.Equal("1", (string)sharedItems.Attribute("containsInteger"));
        Assert.Equal("1", (string)sharedItems.Attribute("containsNonDate"));
        Assert.Equal("0", (string)sharedItems.Attribute("containsString"));
        Assert.Equal("0", (string)sharedItems.Attribute("containsSemiMixedTypes"));
        Assert.Equal("0", (string)sharedItems.Attribute("containsDate"));
        Assert.Equal("0", (string)sharedItems.Attribute("containsBlank"));
        Assert.Equal("0", (string)sharedItems.Attribute("containsMixedTypes"));
        Assert.Equal("0", (string)sharedItems.Attribute("longText"));
    }

    /// <summary>
    /// 验证生产序列化重载同时写入完整缓存、内联列表校验和公式文本。
    /// </summary>
    [Fact]
    public void ContentWriter_ShouldWriteCacheWithInlineValidationAndFormulaText()
    {
        using var workbook = CreateWorkbook(out var pivot);
        var summary = (XSSFSheet)workbook.GetSheet("Summary");
        summary.CreateRow(10).CreateCell(10).SetCellFormula("1+1");
        var records = NpoiPivotCacheWriter.Complete(pivot, Headers, ReadOnlyRows(SourceRows),
            CancellationToken.None);
        var rule = new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 3, StartColumn = 6, EndColumn = 6 },
            Type = ExcelDataValidationType.List,
            Values = new[] { "Open", "Closed" }
        };
        NpoiSheetContentWriter.ApplyValidations(workbook, workbook.GetSheet("Data"), new[] { rule },
            CancellationToken.None);
        IReadOnlyDictionary<string, IReadOnlyList<ExcelDataValidationDefinition>> validations =
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>
            {
                ["Data"] = new[] { rule }
            };
        using var output = new MemoryStream();

        NpoiSheetContentWriter.Write(workbook, validations, new[] { records }, output, CancellationToken.None);

        Assert.True(output.CanWrite);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        var cacheNs = (XNamespace)"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var recordPart = ReadXml(archive, records.PartName);
        Assert.Equal((uint)SourceRows.Length, (uint)recordPart.Root.Elements(cacheNs + "r").Count());
        Assert.Equal(new[] { 6, 6, 6 }, recordPart.Root.Elements(cacheNs + "r")
            .Select(row => row.Elements(cacheNs + "x").Count()));
        var worksheet = ReadXml(archive, "xl/worksheets/sheet1.xml");
        var validation = Assert.Single(worksheet.Descendants(cacheNs + "dataValidation"));
        Assert.Equal("G2:G4", (string)validation.Attribute("sqref"));
        Assert.Equal("\"Open,Closed\"", validation.Element(cacheNs + "formula1").Value);
        var packageRelationshipsNs = (XNamespace)"http://schemas.openxmlformats.org/package/2006/relationships";
        var workbookRelations = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var cacheDefinitionRelation = Assert.Single(workbookRelations.Root
            .Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true));
        var cacheDefinitionPart = ResolvePart("xl/workbook.xml", (string)cacheDefinitionRelation.Attribute("Target"));
        var cacheDefinitionDocument = ReadXml(archive, cacheDefinitionPart);
        var sourceRange = cacheDefinitionDocument.Root.Element(cacheNs + "cacheSource")
            .Element(cacheNs + "worksheetSource");
        Assert.Equal("Data", (string)sourceRange.Attribute("sheet"));
        Assert.Equal("A1:F4", (string)sourceRange.Attribute("ref"));
        var cacheFields = cacheDefinitionDocument.Root.Element(cacheNs + "cacheFields")
            .Elements(cacheNs + "cacheField").ToArray();
        Assert.Equal(Headers, cacheFields.Select(field => (string)field.Attribute("name")));
        Assert.Equal(new[] { 2, 3, 3, 2, 2, 2 }, cacheFields.Select(field =>
            field.Element(cacheNs + "sharedItems").Elements().Count()));
        var cacheRecords = recordPart.Root.Elements(cacheNs + "r").ToArray();
        for (var row = 0; row < SourceRows.Length; row++)
        {
            var indexes = cacheRecords[row].Elements(cacheNs + "x").ToArray();
            for (var column = 0; column < Headers.Length; column++)
            {
                var index = uint.Parse((string)indexes[column].Attribute("v"), CultureInfo.InvariantCulture);
                var item = cacheFields[column].Element(cacheNs + "sharedItems").Elements().ElementAt((int)index);
                AssertSourceValue(SourceRows[row][column], ReadSharedValue(item), $"final-cache[{row},{column}]");
            }
        }
        var workbookXml = ReadXml(archive, "xl/workbook.xml");
        var officeRelationshipNs = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var summarySheet = workbookXml.Root.Element(cacheNs + "sheets").Elements(cacheNs + "sheet")
            .Single(sheet => (string)sheet.Attribute("name") == "Summary");
        var summaryPart = ResolvePart("xl/workbook.xml", (string)workbookRelations.Root
            .Elements(packageRelationshipsNs + "Relationship")
            .Single(element => (string)element.Attribute("Id") == (string)summarySheet.Attribute(officeRelationshipNs + "id"))
            .Attribute("Target"));
        var pivotRelations = ReadXml(archive, RelationshipPath(summaryPart));
        var pivotTableRelation = Assert.Single(pivotRelations.Root.Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotTable",
                StringComparison.Ordinal) == true));
        var pivotPart = ResolvePart(summaryPart, (string)pivotTableRelation.Attribute("Target"));
        var pivotTableRelations = ReadXml(archive, RelationshipPath(pivotPart));
        var pivotCacheRelation = Assert.Single(pivotTableRelations.Root.Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheDefinition",
                StringComparison.Ordinal) == true));
        Assert.Equal(cacheDefinitionPart, ResolvePart(pivotPart, (string)pivotCacheRelation.Attribute("Target")));
        var cacheRelations = ReadXml(archive, RelationshipPath(cacheDefinitionPart));
        var recordsRelation = Assert.Single(cacheRelations.Root.Elements(packageRelationshipsNs + "Relationship")
            .Where(element => ((string)element.Attribute("Type"))?.EndsWith("/pivotCacheRecords",
                StringComparison.Ordinal) == true));
        Assert.Equal(records.PartName, ResolvePart(cacheDefinitionPart, (string)recordsRelation.Attribute("Target")));
        var contentTypes = ReadXml(archive, "[Content_Types].xml");
        Assert.Contains(contentTypes.Root.Elements().Where(element => element.Name.LocalName == "Override"),
            element => (string)element.Attribute("PartName") == "/" + records.PartName);
        using var reopenedInput = new MemoryStream(output.ToArray(), writable: false);
        using var reopened = new XSSFWorkbook(reopenedInput);
        var reopenedData = reopened.GetSheet("Data");
        var reopenedSummary = reopened.GetSheet("Summary");
        for (var column = 0; column < Headers.Length; column++)
            Assert.Equal(Headers[column], reopenedData.GetRow(0).GetCell(column).StringCellValue);
        for (var row = 0; row < SourceRows.Length; row++)
        {
            for (var column = 0; column < Headers.Length; column++)
                AssertSourceValue(SourceRows[row][column], ReadSourceCell(reopenedData.GetRow(row + 1).GetCell(column)),
                    $"reopened-source[{row},{column}]");
        }
        Assert.Equal("1+1", reopenedSummary.GetRow(10).GetCell(10).CellFormula);
    }

    /// <summary>
    /// 验证 ZIP 重写期间取消后保留调用方流所有权。
    /// </summary>
    [Fact]
    public void ContentWriter_CancellationDuringZipRewrite_ShouldKeepCallerStreamOpenAndDeleteStagingFile()
    {
        using var workbook = CreateWorkbook(out var pivot);
        var records = NpoiPivotCacheWriter.Complete(pivot, Headers, ReadOnlyRows(SourceRows),
            CancellationToken.None);
        using var cancellation = new CancellationTokenSource();
        using var output = new CancelAfterFirstWriteStream(cancellation);
        var tempDirectory = Path.GetTempPath();
        var stagedFilesBefore = new HashSet<string>(
            Directory.GetFiles(tempDirectory, "bing-npoi-validations-*.tmp"), StringComparer.OrdinalIgnoreCase);
        IReadOnlyDictionary<string, IReadOnlyList<ExcelDataValidationDefinition>> validations =
            new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>();

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiSheetContentWriter.Write(workbook, validations, new[] { records }, output, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.True(output.Length > 0);
        Assert.True(output.CanWrite);
        Assert.False(output.WasDisposed);
        Assert.Empty(Directory.GetFiles(tempDirectory, "bing-npoi-validations-*.tmp")
            .Where(path => !stagedFilesBefore.Contains(path)));
    }

    /// <summary>
    /// 验证缓存记录在完成时隔离输入集合并支持相同来源的重复导出。
    /// </summary>
    [Fact]
    public void Complete_ShouldSnapshotInputAndRemainRepeatable()
    {
        var mutableRows = SourceRows.Select(row => (object[])row.Clone()).ToArray();
        using var firstWorkbook = CreateWorkbook(out var firstPivot);
        var firstRecords = NpoiPivotCacheWriter.Complete(firstPivot, Headers, ReadOnlyRows(mutableRows),
            CancellationToken.None);
        mutableRows[0][0] = "Changed";
        mutableRows[0][1] = 999d;

        Assert.Equal(new uint[] { 0, 0, 0, 0, 0, 0 }, firstRecords.Rows[0]);
        Assert.Equal("North", Assert.IsType<CT_String>(firstPivot.GetPivotCacheDefinition()
            .GetCTPivotCacheDefinition().cacheFields.cacheField[0].sharedItems.Items[0]).v);
        Assert.Equal(10d, Assert.IsType<CT_Number>(firstPivot.GetPivotCacheDefinition()
            .GetCTPivotCacheDefinition().cacheFields.cacheField[1].sharedItems.Items[0]).v);

        using var secondWorkbook = CreateWorkbook(out var secondPivot);
        var secondRecords = NpoiPivotCacheWriter.Complete(secondPivot, Headers, ReadOnlyRows(SourceRows),
            CancellationToken.None);
        Assert.Equal(firstRecords.PartName, secondRecords.PartName);
        Assert.Equal(firstRecords.SharedItemCounts, secondRecords.SharedItemCounts);
        Assert.Equal(firstRecords.Rows.Select(row => row.ToArray()), secondRecords.Rows.Select(row => row.ToArray()));
    }

    /// <summary>
    /// 验证预取消在缓存替换前传播并保留原有缓存定义。
    /// </summary>
    [Fact]
    public void Complete_PreCanceled_ShouldLeaveCacheDefinitionUnchanged()
    {
        using var workbook = CreateWorkbook(out var pivot);
        var definition = pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition();
        var originalFields = definition.cacheFields;
        var fields = originalFields.cacheField;
        var originalNames = fields.Select(field => field.name).ToArray();
        var originalRecordCount = definition.recordCount;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotCacheWriter.Complete(pivot, Headers, ReadOnlyRows(SourceRows), cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Same(originalFields, definition.cacheFields);
        Assert.Equal(originalRecordCount, definition.recordCount);
        Assert.Equal(originalNames, fields.Select(field => field.name));
        Assert.All(fields, field => Assert.Empty(field.sharedItems.Items));
    }

    /// <summary>
    /// 验证缓存字段构建中取消时不替换原有缓存定义。
    /// </summary>
    [Fact]
    public void Complete_CanceledDuringCacheBuild_ShouldLeaveCacheDefinitionUnchanged()
    {
        using var workbook = CreateWorkbook(out var pivot);
        var definition = pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition();
        var originalFields = definition.cacheFields;
        var fields = originalFields.cacheField;
        var originalNames = fields.Select(field => field.name).ToArray();
        var originalRecordCount = definition.recordCount;
        using var cancellation = new CancellationTokenSource();
        var rows = new CancelOnReadRows(ReadOnlyRows(SourceRows), cancellation,
            SourceRows.Length * 2 + 1);

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotCacheWriter.Complete(pivot, Headers, rows, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Same(originalFields, definition.cacheFields);
        Assert.Equal(originalRecordCount, definition.recordCount);
        Assert.Equal(originalNames, fields.Select(field => field.name));
        Assert.All(fields, field => Assert.Empty(field.sharedItems.Items));
    }

    /// <summary>
    /// 验证错误行宽和非有限数值在缓存定义变更前被拒绝。
    /// </summary>
    /// <param name="caseName">触发校验失败的输入矩阵类别。</param>
    [Theory]
    [InlineData("wrong-width")]
    [InlineData("nan")]
    [InlineData("positive-infinity")]
    public void Complete_InvalidSource_ShouldLeaveCacheDefinitionUnchanged(string caseName)
    {
        using var workbook = CreateWorkbook(out var pivot);
        var definition = pivot.GetPivotCacheDefinition().GetCTPivotCacheDefinition();
        var originalFields = definition.cacheFields;
        var fields = originalFields.cacheField;
        var originalNames = fields.Select(field => field.name).ToArray();
        var originalRecordCount = definition.recordCount;

        Assert.Throws<ArgumentException>(() => NpoiPivotCacheWriter.Complete(pivot, Headers,
            InvalidRows(caseName), CancellationToken.None));

        Assert.Same(originalFields, definition.cacheFields);
        Assert.Equal(originalRecordCount, definition.recordCount);
        Assert.Equal(originalNames, fields.Select(field => field.name));
        Assert.All(fields, field => Assert.Empty(field.sharedItems.Items));
    }

    /// <summary>
    /// 验证缓存记录快照拒绝错误列数和越界共享项索引。
    /// </summary>
    [Fact]
    public void RecordsConstructor_ShouldRejectInvalidWidthsAndIndexes()
    {
        Assert.Throws<ArgumentException>(() => new NpoiPivotCacheRecords("xl/pivotCache/records.xml",
            new[] { 1, 1 }, new IReadOnlyList<uint>[] { new uint[] { 0 } }));
        Assert.Throws<ArgumentException>(() => new NpoiPivotCacheRecords("xl/pivotCache/records.xml",
            new[] { 1, 1 }, new IReadOnlyList<uint>[] { new uint[] { 0, 1 } }));
    }

    /// <summary>
    /// 验证记录构造器复制输入并公开不可写的集合快照。
    /// </summary>
    [Fact]
    public void RecordsConstructor_ShouldSnapshotCountsAndRows()
    {
        var counts = new[] { 2, 2 };
        var row = new uint[] { 1, 0 };
        var rows = new IReadOnlyList<uint>[] { row };
        var records = new NpoiPivotCacheRecords("/xl/pivotCache/records.xml", counts, rows);
        counts[0] = 9;
        row[0] = 0;
        rows[0] = new uint[] { 0, 0 };

        Assert.Equal("xl/pivotCache/records.xml", records.PartName);
        Assert.Equal(new[] { 2, 2 }, records.SharedItemCounts);
        Assert.Equal(new uint[] { 1, 0 }, records.Rows[0]);
        Assert.Throws<NotSupportedException>(() => ((IList<int>)records.SharedItemCounts)[0] = 5);
        Assert.Throws<NotSupportedException>(() => ((IList<uint>)records.Rows[0])[0] = 0);
    }

    /// <summary>
    /// 验证记录写入取消传播令牌且不关闭调用方流。
    /// </summary>
    [Fact]
    public void RecordWriter_Cancellation_ShouldKeepCallerStreamOpen()
    {
        var records = new NpoiPivotCacheRecords("xl/pivotCache/records.xml", new[] { 2, 2 },
            new IReadOnlyList<uint>[] { new uint[] { 0, 1 }, new uint[] { 1, 0 } });
        using var cancellation = new CancellationTokenSource();
        using var output = new CancelOnFirstWriteStream(cancellation);

        var exception = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotCacheRecordWriter.Write(output, records, cancellation.Token));

        Assert.True(output.BytesForwarded > 0);
        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.False(output.WasDisposed);

        using var preCanceledSource = new CancellationTokenSource();
        preCanceledSource.Cancel();
        using var untouchedOutput = new MemoryStream();
        var preCanceled = Assert.Throws<OperationCanceledException>(() =>
            NpoiPivotCacheRecordWriter.Write(untouchedOutput, records, preCanceledSource.Token));
        Assert.Equal(preCanceledSource.Token, preCanceled.CancellationToken);
        Assert.Empty(untouchedOutput.ToArray());
        Assert.True(untouchedOutput.CanWrite);
    }

    /// <summary>
    /// 验证缓存写入要求 XLSX，并在缺少记录部件时保留目标流。
    /// </summary>
    [Fact]
    public void ContentWriter_ShouldRejectHssfAndMissingRecordPartBeforeDestinationWrite()
    {
        using var hssf = new NPOI.HSSF.UserModel.HSSFWorkbook();
        using var hssfOutput = new MemoryStream();
        var cacheRecords = new NpoiPivotCacheRecords("xl/pivotCache/pivotCacheRecords1.xml",
            new[] { 1 }, new IReadOnlyList<uint>[] { new uint[] { 0 } });
        var validations = new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>();

        Assert.Throws<ArgumentException>(() => NpoiSheetContentWriter.Write(hssf, validations,
            new[] { cacheRecords }, hssfOutput, CancellationToken.None));

        Assert.Empty(hssfOutput.ToArray());
        Assert.True(hssfOutput.CanWrite);

        using var xssf = CreateWorkbook(out _);
        using var xssfOutput = new MemoryStream();
        var missingPart = new NpoiPivotCacheRecords("xl/pivotCache/not-present.xml", new[] { 1 },
            new IReadOnlyList<uint>[] { new uint[] { 0 } });
        var missing = Assert.Throws<BingOfficesConfigurationException>(() =>
            NpoiSheetContentWriter.Write(xssf, validations, new[] { missingPart }, xssfOutput,
                CancellationToken.None));
        Assert.Equal(BingOfficesStage.Write, missing.Stage);
        Assert.Empty(xssfOutput.ToArray());
        Assert.True(xssfOutput.CanWrite);
    }

    /// <summary>
    /// 验证没有透视缓存时原有 XLS/HSSF 与 XLSX/XSSF 工作簿写入保持可用。
    /// </summary>
    [Fact]
    public void ContentWriter_WithoutPivotRecords_ShouldKeepXlsAndXlsxOutput()
    {
        var validations = new Dictionary<string, IReadOnlyList<ExcelDataValidationDefinition>>();
        using var hssf = new NPOI.HSSF.UserModel.HSSFWorkbook();
        hssf.CreateSheet("Rows").CreateRow(0).CreateCell(0).SetCellValue("HSSF");
        using var hssfOutput = new MemoryStream();
        NpoiSheetContentWriter.Write(hssf, validations,
            new NpoiNonDisposingStream(hssfOutput, CancellationToken.None), CancellationToken.None);
        Assert.True(hssfOutput.CanWrite);
        hssfOutput.Position = 0;
        using (var reopened = WorkbookFactory.Create(hssfOutput))
            Assert.Equal("HSSF", reopened.GetSheet("Rows").GetRow(0).GetCell(0).StringCellValue);

        using var xssf = new XSSFWorkbook();
        xssf.CreateSheet("Rows").CreateRow(0).CreateCell(0).SetCellValue("XSSF");
        using var xssfOutput = new MemoryStream();
        NpoiSheetContentWriter.Write(xssf, validations,
            new NpoiNonDisposingStream(xssfOutput, CancellationToken.None), CancellationToken.None);
        Assert.True(xssfOutput.CanWrite);
        xssfOutput.Position = 0;
        using var xssfReopened = new XSSFWorkbook(xssfOutput);
        Assert.Equal("XSSF", xssfReopened.GetSheet("Rows").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 创建含真实源数据和原生字段角色的 XSSF 透视表。
    /// </summary>
    /// <param name="pivot">返回接收缓存定义的透视表。</param>
    /// <returns>源工作簿。</returns>
    private static XSSFWorkbook CreateWorkbook(out XSSFPivotTable pivot)
    {
        var workbook = new XSSFWorkbook();
        var data = workbook.CreateSheet("Data");
        var summary = (XSSFSheet)workbook.CreateSheet("Summary");
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");
        var header = data.CreateRow(0);
        for (var column = 0; column < Headers.Length; column++)
            header.CreateCell(column, CellType.String).SetCellValue(Headers[column]);
        for (var rowIndex = 0; rowIndex < SourceRows.Length; rowIndex++)
        {
            var row = data.CreateRow(rowIndex + 1);
            for (var column = 0; column < Headers.Length; column++)
            {
                var cell = row.CreateCell(column);
                switch (SourceRows[rowIndex][column])
                {
                    case null:
                        cell.SetCellType(CellType.Blank);
                        break;
                    case string text:
                        cell.SetCellType(CellType.String);
                        cell.SetCellValue(text);
                        break;
                    case double number:
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(number);
                        break;
                    case bool boolean:
                        cell.SetCellType(CellType.Boolean);
                        cell.SetCellValue(boolean);
                        break;
                    case DateTime date:
                        cell.SetCellType(CellType.Numeric);
                        cell.SetCellValue(date);
                        cell.CellStyle = dateStyle;
                        break;
                    default:
                        throw new InvalidOperationException("Unexpected source cell type.");
                }
            }
        }

        var source = new AreaReference(new CellReference("A1"),
            new CellReference(SourceRows.Length, Headers.Length - 1), SpreadsheetVersion.EXCEL2007);
        pivot = summary.CreatePivotTable(source, new CellReference("A1"), data);
        pivot.AddRowLabel(0);
        pivot.AddColumnLabel(DataConsolidateFunction.SUM, 1, "Total");
        return workbook;
    }

    /// <summary>
    /// 创建单字段混合类型来源及对应原生行轴透视表。
    /// </summary>
    /// <param name="headerText">唯一来源字段标题。</param>
    /// <param name="values">按原始行顺序排列的值。</param>
    /// <param name="pivot">返回接收缓存定义的透视表。</param>
    /// <returns>包含真实单元格和原生透视表的工作簿。</returns>
    private static XSSFWorkbook CreateSingleFieldWorkbook(string headerText, object[] values,
        out XSSFPivotTable pivot)
    {
        var workbook = new XSSFWorkbook();
        var data = workbook.CreateSheet("Data");
        var summary = (XSSFSheet)workbook.CreateSheet("Summary");
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");
        data.CreateRow(0).CreateCell(0, CellType.String).SetCellValue(headerText);
        for (var rowIndex = 0; rowIndex < values.Length; rowIndex++)
        {
            var cell = data.CreateRow(rowIndex + 1).CreateCell(0);
            switch (values[rowIndex])
            {
                case null:
                    cell.SetCellType(CellType.Blank);
                    break;
                case string text:
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(text);
                    break;
                case double number:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(number);
                    break;
                case bool boolean:
                    cell.SetCellType(CellType.Boolean);
                    cell.SetCellValue(boolean);
                    break;
                case DateTime date:
                    cell.SetCellType(CellType.Numeric);
                    cell.SetCellValue(date);
                    cell.CellStyle = dateStyle;
                    break;
                default:
                    throw new InvalidOperationException("Unexpected mixed source cell type.");
            }
        }
        var source = new AreaReference(new CellReference("A1"),
            new CellReference(values.Length, 0), SpreadsheetVersion.EXCEL2007);
        pivot = summary.CreatePivotTable(source, new CellReference("A1"), data);
        pivot.AddRowLabel(0);
        return workbook;
    }

    /// <summary>
    /// 将来源行数组包装为缓存 API 要求的只读二维集合。
    /// </summary>
    /// <param name="rows">来源记录。</param>
    /// <returns>只读来源记录集合。</returns>
    private static IReadOnlyList<IReadOnlyList<object>> ReadOnlyRows(object[][] rows) =>
        rows.Select(row => (IReadOnlyList<object>)row).ToArray();

    /// <summary>
    /// 创建错误行宽或含非有限数字的输入矩阵。
    /// </summary>
    /// <param name="caseName">错误输入类别。</param>
    /// <returns>测试使用的来源行。</returns>
    private static IReadOnlyList<IReadOnlyList<object>> InvalidRows(string caseName)
    {
        var rows = SourceRows.Select(row => (object[])row.Clone()).ToArray();
        switch (caseName)
        {
            case "wrong-width":
                return rows.Select(row => (IReadOnlyList<object>)row.Take(Headers.Length - 1).ToArray()).ToArray();
            case "nan":
                rows[0][1] = double.NaN;
                break;
            case "positive-infinity":
                rows[1][1] = double.PositiveInfinity;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(caseName));
        }
        return ReadOnlyRows(rows);
    }

    /// <summary>
    /// 按关系部件路径读取 OOXML 文档。
    /// </summary>
    /// <param name="archive">工作簿压缩包。</param>
    /// <param name="partName">包内部件路径。</param>
    /// <returns>解析后的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string partName)
    {
        var entry = archive.GetEntry(partName)
            ?? throw new InvalidDataException($"Missing OOXML part {partName}.");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 还原关系中的相对包目标路径。
    /// </summary>
    /// <param name="sourcePart">关系所在源部件。</param>
    /// <param name="target">关系目标。</param>
    /// <returns>规范化的包内部件路径。</returns>
    private static string ResolvePart(string sourcePart, string target)
    {
        if (target.StartsWith("/", StringComparison.Ordinal))
            return target.TrimStart('/');
        return Uri.UnescapeDataString(new Uri(new Uri("https://ooxml.invalid/" + sourcePart), target)
            .AbsolutePath.TrimStart('/'));
    }

    /// <summary>
    /// 获取指定部件的关系部件路径。
    /// </summary>
    /// <param name="partPath">包内源部件路径。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPath(string partPath)
    {
        var directory = Path.GetDirectoryName(partPath.Replace('/', Path.DirectorySeparatorChar)) ?? string.Empty;
        return Path.Combine(directory, "_rels", Path.GetFileName(partPath) + ".rels")
            .Replace(Path.DirectorySeparatorChar, '/');
    }

    /// <summary>
    /// 将缓存共享项解码为 CLR 标量。
    /// </summary>
    /// <param name="item">缓存共享项 XML 元素。</param>
    /// <returns>文本、数字、布尔、日期或空值。</returns>
    private static object ReadSharedValue(XElement item)
    {
        var value = (string)item.Attribute("v");
        return item.Name.LocalName switch
        {
            "m" => null,
            "s" => DecodeXstring(value ?? string.Empty),
            "n" => double.Parse(value, CultureInfo.InvariantCulture),
            "b" => value == "1",
            "d" => DateTime.Parse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind),
            _ => throw new InvalidDataException($"Unexpected cache item type {item.Name.LocalName}.")
        };
    }

    /// <summary>
    /// 解码 SpreadsheetML 文本属性中的一次 Xstring 转义。
    /// </summary>
    /// <param name="value">缓存定义中的编码文本。</param>
    /// <returns>还原一次转义后的原始文本。</returns>
    private static string DecodeXstring(string value)
    {
        var decoded = new System.Text.StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            if (index + 6 < value.Length && value[index] == '_' && value[index + 1] == 'x'
                && value[index + 6] == '_' && Uri.IsHexDigit(value[index + 2])
                && Uri.IsHexDigit(value[index + 3]) && Uri.IsHexDigit(value[index + 4])
                && Uri.IsHexDigit(value[index + 5]))
            {
                decoded.Append((char)int.Parse(value.Substring(index + 2, 4), NumberStyles.HexNumber,
                    CultureInfo.InvariantCulture));
                index += 6;
            }
            else
                decoded.Append(value[index]);
        }
        return decoded.ToString();
    }

    /// <summary>
    /// 读取 OOXML 中可省略的布尔标志属性。
    /// </summary>
    /// <param name="element">包含缓存标志的 XML 元素。</param>
    /// <param name="attributeName">待读取的属性名称。</param>
    /// <returns>属性值；未序列化时返回 false。</returns>
    private static bool XmlFlag(XElement element, string attributeName)
    {
        var document = new XmlDocument();
        document.LoadXml(element.ToString(SaveOptions.DisableFormatting));
        var namespaces = new XmlNamespaceManager(document.NameTable);
        namespaces.AddNamespace("x", element.Name.NamespaceName);
        var parsed = CT_SharedItems.Parse(document.DocumentElement, namespaces);
        return attributeName switch
        {
            "containsString" => parsed.containsString,
            "containsMixedTypes" => parsed.containsMixedTypes,
            "containsSemiMixedTypes" => parsed.containsSemiMixedTypes,
            "containsNumber" => parsed.containsNumber,
            "containsDate" => parsed.containsDate,
            "containsNonDate" => parsed.containsNonDate,
            "containsBlank" => parsed.containsBlank,
            "containsInteger" => parsed.containsInteger,
            "longText" => parsed.longText,
            _ => throw new ArgumentOutOfRangeException(nameof(attributeName))
        };
    }

    /// <summary>
    /// 读取重开工作簿中的源单元格 CLR 值。
    /// </summary>
    /// <param name="cell">源工作表单元格。</param>
    /// <returns>文本、数字、布尔、日期或空值。</returns>
    private static object ReadSourceCell(ICell cell)
    {
        if (cell == null || cell.CellType == CellType.Blank)
            return null;
        return cell.CellType switch
        {
            CellType.String => cell.StringCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => cell.DateCellValue,
            CellType.Numeric => cell.NumericCellValue,
            _ => throw new InvalidDataException($"Unexpected source cell type {cell.CellType}.")
        };
    }

    /// <summary>
    /// 比较缓存或源单元格的值和 CLR 类型。
    /// </summary>
    /// <param name="expected">来源预期值。</param>
    /// <param name="actual">实际读取值。</param>
    /// <param name="location">待比较的缓存字段坐标。</param>
    private static void AssertSourceValue(object expected, object actual, string location)
    {
        Assert.Equal(expected?.GetType(), actual?.GetType());
        if (expected is DateTime expectedDate && actual is DateTime actualDate)
        {
            Assert.Equal(expectedDate, actualDate);
            return;
        }
        Assert.True(Equals(expected, actual),
            $"Value mismatch at {location}: expected '{expected}', actual '{actual}'.");
    }

    /// <summary>
    /// 在读取指定集合项后触发取消的只读行集合。
    /// </summary>
    private sealed class CancelOnReadRows : IReadOnlyList<IReadOnlyList<object>>
    {
        /// <summary>
        /// 被代理的来源行集合。
        /// </summary>
        private readonly IReadOnlyList<IReadOnlyList<object>> _rows;

        /// <summary>
        /// 达到指定读取次数时触发的取消令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 触发取消时对应的集合读取次数。
        /// </summary>
        private readonly int _cancelOnRead;

        /// <summary>
        /// 已读取的集合项次数。
        /// </summary>
        private int _readCount;

        /// <summary>
        /// 初始化一个 <see cref="CancelOnReadRows" /> 类型的实例。
        /// </summary>
        /// <param name="rows">被代理的来源行集合。</param>
        /// <param name="cancellation">在指定读取次数触发的令牌源。</param>
        /// <param name="cancelOnRead">触发取消的读取次数。</param>
        internal CancelOnReadRows(IReadOnlyList<IReadOnlyList<object>> rows,
            CancellationTokenSource cancellation, int cancelOnRead)
        {
            _rows = rows;
            _cancellation = cancellation;
            _cancelOnRead = cancelOnRead;
        }

        /// <inheritdoc />
        public int Count => _rows.Count;

        /// <inheritdoc />
        public IReadOnlyList<object> this[int index]
        {
            get
            {
                var row = _rows[index];
                if (++_readCount == _cancelOnRead)
                    _cancellation.Cancel();
                return row;
            }
        }

        /// <inheritdoc />
        public IEnumerator<IReadOnlyList<object>> GetEnumerator() => _rows.GetEnumerator();

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// 在首次写入数据后请求取消的内存流。
    /// </summary>
    private sealed class CancelAfterFirstWriteStream : MemoryStream
    {
        /// <summary>
        /// 首次写入后用于请求取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 标记是否已请求取消。
        /// </summary>
        private bool _canceled;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">首次写入后触发的令牌源。</param>
        internal CancelAfterFirstWriteStream(CancellationTokenSource cancellation)
        {
            _cancellation = cancellation ?? throw new ArgumentNullException(nameof(cancellation));
        }

        /// <summary>
        /// 获取流是否已释放。
        /// </summary>
        internal bool WasDisposed { get; private set; }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            base.Write(buffer, offset, count);
            CancelAfterFirstWrite();
        }

        /// <inheritdoc />
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            base.Write(buffer);
            CancelAfterFirstWrite();
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            WasDisposed = true;
            base.Dispose(disposing);
        }

        /// <summary>
        /// 在首次写入后请求取消。
        /// </summary>
        private void CancelAfterFirstWrite()
        {
            if (_canceled)
                return;
            _canceled = true;
            _cancellation.Cancel();
        }
    }
}
