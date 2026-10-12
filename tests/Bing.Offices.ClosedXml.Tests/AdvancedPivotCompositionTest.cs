using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Tests.AdvancedPivotComposition;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证高级透视与模板、字典、保护和表格组合时的隔离与复用。
/// </summary>
public sealed class AdvancedPivotCompositionTest
{
    /// <summary>
    /// SpreadsheetML 工作簿和透视部件使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Office Open XML 文档间关系使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace OfficeRelationshipNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Open Packaging Conventions 部件关系使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace PackageRelationshipNs =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// 验证同一来源的不同透视配置及重复导出保持隔离。
    /// </summary>
    [Fact]
    public void SameSourcePivots_ShouldRemainIsolatedAcrossRepeatedExports()
    {
        using var template = CreateTemplate();
        var request = CreateRequest(template);
        var exporter = new ClosedXmlExcelExporter();
        using var first = new MemoryStream();
        using var second = new MemoryStream();

        exporter.Export(request, first);
        exporter.Export(request, second);

        AssertComposition(first.ToArray());
        AssertComposition(second.ToArray());
        var firstPackage = ReadPackage(first.ToArray());
        var secondPackage = ReadPackage(second.ToArray());
        Assert.True(XNode.DeepEquals(GetPivotByName(firstPackage, "ByRegionChannel"),
            GetPivotByName(secondPackage, "ByRegionChannel")));
        Assert.True(XNode.DeepEquals(GetPivotByName(firstPackage, "ByStatus"),
            GetPivotByName(secondPackage, "ByStatus")));
    }

    /// <summary>
    /// 验证预取消失败后，同一请求仍可成功导出完整组合工作簿。
    /// </summary>
    [Fact]
    public async Task FailedExport_ShouldAllowSameRequestReuse()
    {
        using var template = CreateTemplate();
        var request = CreateRequest(template);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var original = new byte[] { 3, 5, 8, 13 };
        using var failedOutput = new MemoryStream();
        failedOutput.Write(original, 0, original.Length);
        failedOutput.Position = 0;
        var exporter = new ClosedXmlExcelExporter();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            exporter.ExportAsync(request, failedOutput, cancellation.Token));

        Assert.Equal(original, failedOutput.ToArray());
        Assert.True(failedOutput.CanWrite);
        using var successfulOutput = new MemoryStream();
        await exporter.ExportAsync(request, successfulOutput);
        AssertComposition(successfulOutput.ToArray());
    }

    /// <summary>
    /// 创建包含普通内容、可保留的隐藏工作表和字典名称冲突项的模板。
    /// </summary>
    /// <returns>定位到起始位置的 XLSX 模板流。</returns>
    private static MemoryStream CreateTemplate()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var data = workbook.Worksheets.Add("Data");
            data.Cell("F1").Value = "DataTemplateMarker";
            var entry = workbook.Worksheets.Add("Entry");
            entry.Cell("F1").Value = "EntryTemplateMarker";
            entry.Cell("F1").Style.Fill.BackgroundColor = XLColor.Yellow;
            var archive = workbook.Worksheets.Add("Archive");
            archive.Cell("A1").Value = "ArchivedTemplateValue";
            archive.Visibility = XLWorksheetVisibility.Hidden;
            workbook.Worksheets.Add("Notes").Cell("A1").Value = "OrdinaryTemplateValue";
            var occupiedLookup = workbook.Worksheets.Add("__BingLookup_001");
            occupiedLookup.Cell("A1").Value = "ExistingLookupSheet";
            occupiedLookup.Visibility = XLWorksheetVisibility.Hidden;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建同源双透视、命名字典、保护区和数据表的组合请求。
    /// </summary>
    /// <param name="template">包含普通与隐藏模板工作表的流。</param>
    /// <returns>已快照的组合工作簿请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template) =>
        ExcelExport.Workbook(builder => builder
            .UseTemplate(template, leaveOpen: true)
            .LookupList("OrderStatus", new[] { "Open", "Closed" })
            .AddSheet("Data", DataRows(), sheet => sheet.Table(new ExcelTableDefinition
            {
                Name = "CompositionData",
                Range = new ExcelRangeDefinition
                {
                    StartRow = 0, StartColumn = 0, EndRow = 4, EndColumn = 4
                },
                StyleName = "TableStyleMedium2"
            }))
            .AddSheet("Entry", new[]
            {
                new AdvancedPivotCompositionEntry { OrderId = "PO-1", Status = "Open" },
                new AdvancedPivotCompositionEntry { OrderId = "PO-2", Status = "Closed" }
            }, sheet => sheet
                .DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 2, StartColumn = 1, EndColumn = 1
                    },
                    Type = ExcelDataValidationType.List,
                    ListSourceName = "OrderStatus"
                })
                .Protection(new ExcelSheetProtectionDefinition(new[]
                {
                    new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 1
                    }
                }, "entry-password")))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "ByRegionChannel",
                SourceSheet = "Data",
                SourceRange = SourceRange(),
                TargetSheet = "RegionSummary",
                TargetRow = 1,
                TargetColumn = 2,
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition
                    {
                        Field = "Region", SortOrder = ExcelPivotSortOrder.Ascending
                    }
                },
                ColumnFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition
                    {
                        Field = "Channel", SortOrder = ExcelPivotSortOrder.Descending,
                        SubtotalMode = ExcelPivotSubtotalMode.None
                    }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition
                    {
                        Field = "Amount", Caption = "Open total", NumberFormat = "0.00"
                    },
                    new ExcelPivotValueDefinition
                    {
                        Field = "OrderId", Caption = "Open order count",
                        Aggregation = ExcelPivotAggregation.Count
                    }
                },
                Filters = new[]
                {
                    new ExcelPivotFilterDefinition
                    {
                        Field = "Status", Values = new object[] { "Open" }
                    }
                }
            })
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "ByStatus",
                SourceSheet = "Data",
                SourceRange = SourceRange(),
                TargetSheet = "StatusSummary",
                TargetRow = 2,
                TargetColumn = 1,
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition
                    {
                        Field = "Status", SortOrder = ExcelPivotSortOrder.Descending,
                        SubtotalMode = ExcelPivotSubtotalMode.None
                    }
                },
                ColumnFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition
                    {
                        Field = "Region", SortOrder = ExcelPivotSortOrder.Ascending
                    }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition
                    {
                        Field = "Amount", Caption = "Retail average",
                        Aggregation = ExcelPivotAggregation.Average, NumberFormat = "0.0"
                    },
                    new ExcelPivotValueDefinition
                    {
                        Field = "OrderId", Caption = "Retail order count",
                        Aggregation = ExcelPivotAggregation.Count
                    }
                },
                Filters = new[]
                {
                    new ExcelPivotFilterDefinition
                    {
                        Field = "Channel", Values = new object[] { "Retail" }
                    }
                }
            }));

    /// <summary>
    /// 创建五行业务来源，供两套互不相同的透视配置共享。
    /// </summary>
    /// <returns>完整的采购订单来源数据。</returns>
    private static IEnumerable<AdvancedPivotCompositionRow> DataRows()
    {
        yield return new AdvancedPivotCompositionRow
            { OrderId = "PO-1", Region = "North", Status = "Open", Channel = "Retail", Amount = 10 };
        yield return new AdvancedPivotCompositionRow
            { OrderId = "PO-2", Region = "South", Status = "Closed", Channel = "Online", Amount = 20 };
        yield return new AdvancedPivotCompositionRow
            { OrderId = "PO-3", Region = "North", Status = "Open", Channel = "Online", Amount = 30 };
        yield return new AdvancedPivotCompositionRow
            { OrderId = "PO-4", Region = "South", Status = "Open", Channel = "Retail", Amount = 15 };
    }

    /// <summary>
    /// 返回包含标题行和四条业务记录的零基来源区域。
    /// </summary>
    /// <returns>覆盖 Data 工作表 A1:E5 的来源区域。</returns>
    private static ExcelRangeDefinition SourceRange() => new()
    {
        StartRow = 0, StartColumn = 0, EndRow = 4, EndColumn = 4
    };

    /// <summary>
    /// 验证组合工作簿内容、名称、保护、表格、透视配置、缓存与关系。
    /// </summary>
    /// <param name="bytes">导出的完整 XLSX 字节。</param>
    private static void AssertComposition(byte[] bytes)
    {
        using (var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false)))
        {
            Assert.Equal(new[]
            {
                "Data", "Entry", "Archive", "Notes", "__BingLookup_001", "__BingLookup_002",
                "RegionSummary", "StatusSummary"
            }, workbook.Worksheets.Select(sheet => sheet.Name));
            var data = workbook.Worksheet("Data");
            Assert.Equal(new[]
            {
                "OrderId", "Region", "Status", "Channel", "Amount",
                "PO-1", "North", "Open", "Retail", "10",
                "PO-2", "South", "Closed", "Online", "20",
                "PO-3", "North", "Open", "Online", "30",
                "PO-4", "South", "Open", "Retail", "15"
            }, data.Range("A1:E5").Cells().Select(cell => cell.GetString()));
            Assert.Equal("DataTemplateMarker", data.Cell("F1").GetString());
            Assert.True(data.Tables.Contains("CompositionData"));
            var tableRange = data.Tables.Table("CompositionData").RangeAddress;
            Assert.Equal((1, 1, 5, 5), (tableRange.FirstAddress.RowNumber,
                tableRange.FirstAddress.ColumnNumber, tableRange.LastAddress.RowNumber,
                tableRange.LastAddress.ColumnNumber));

            var entry = workbook.Worksheet("Entry");
            Assert.Equal(new[] { "OrderId", "Status", "PO-1", "Open", "PO-2", "Closed" },
                entry.Range("A1:B3").Cells().Select(cell => cell.GetString()));
            Assert.True(entry.Protection.IsProtected);
            Assert.True(entry.Protection.IsPasswordProtected);
            Assert.False(entry.Cell("A2").Style.Protection.Locked);
            Assert.False(entry.Cell("B3").Style.Protection.Locked);
            Assert.True(entry.Cell("C2").Style.Protection.Locked);
            Assert.Equal("EntryTemplateMarker", entry.Cell("F1").GetString());
            Assert.Equal(XLColor.Yellow, entry.Cell("F1").Style.Fill.BackgroundColor);
            var validation = Assert.Single(entry.DataValidations.GetAllInRange(
                entry.Range("B2:B3").RangeAddress));
            Assert.Equal(XLAllowedValues.List, validation.AllowedValues);
            Assert.Equal("OrderStatus", validation.Value.TrimStart('='));

            Assert.Equal(XLWorksheetVisibility.Hidden, workbook.Worksheet("Archive").Visibility);
            Assert.Equal("ArchivedTemplateValue", workbook.Worksheet("Archive").Cell("A1").GetString());
            Assert.Equal("OrdinaryTemplateValue", workbook.Worksheet("Notes").Cell("A1").GetString());
            Assert.Equal(XLWorksheetVisibility.Hidden, workbook.Worksheet("__BingLookup_001").Visibility);
            Assert.Equal("ExistingLookupSheet", workbook.Worksheet("__BingLookup_001").Cell("A1").GetString());
            Assert.Equal(XLWorksheetVisibility.Hidden, workbook.Worksheet("__BingLookup_002").Visibility);
            Assert.Equal(new[] { "Open", "Closed" },
                new[] { workbook.Worksheet("__BingLookup_002").Cell("A1").GetString(),
                    workbook.Worksheet("__BingLookup_002").Cell("A2").GetString() });
            var definedName = Assert.Single(workbook.DefinedNames.ValidNamedRanges()
                .Where(item => item.Name == "OrderStatus"));
            var nameRange = Assert.Single(definedName.Ranges);
            Assert.Equal("__BingLookup_002", nameRange.Worksheet.Name);
            Assert.Equal((1, 1, 2, 1), (nameRange.RangeAddress.FirstAddress.RowNumber,
                nameRange.RangeAddress.FirstAddress.ColumnNumber, nameRange.RangeAddress.LastAddress.RowNumber,
                nameRange.RangeAddress.LastAddress.ColumnNumber));

            var firstPivot = Assert.Single(workbook.Worksheet("RegionSummary").PivotTables);
            Assert.Equal("ByRegionChannel", firstPivot.Name);
            Assert.Equal(XLPivotSortType.Ascending, firstPivot.RowLabels.Get("Region").SortType);
            Assert.Equal(XLPivotSortType.Descending, firstPivot.ColumnLabels.Get("Channel").SortType);
            Assert.Empty(firstPivot.ColumnLabels.Get("Channel").Subtotals);
            var secondPivot = Assert.Single(workbook.Worksheet("StatusSummary").PivotTables);
            Assert.Equal("ByStatus", secondPivot.Name);
            Assert.Equal(XLPivotSortType.Descending, secondPivot.RowLabels.Get("Status").SortType);
            Assert.Empty(secondPivot.RowLabels.Get("Status").Subtotals);
            Assert.Equal(XLPivotSortType.Ascending, secondPivot.ColumnLabels.Get("Region").SortType);
        }

        var package = ReadPackage(bytes);
        var expectedHeaders = new[] { "OrderId", "Region", "Status", "Channel", "Amount" };
        var rows = new[]
        {
            "PO-1|North|Open|Retail|10", "PO-2|South|Closed|Online|20",
            "PO-3|North|Open|Online|30", "PO-4|South|Open|Retail|15"
        };
        AssertPivot(package, "RegionSummary", "ByRegionChannel", "A1:E5", expectedHeaders, rows,
            new[] { "Region" }, new[] { "Channel", "Values" }, new[] { "Amount|Open total|sum|0.00",
                "OrderId|Open order count|count|" }, "Status", "Open");
        AssertPivot(package, "StatusSummary", "ByStatus", "A1:E5", expectedHeaders, rows,
            new[] { "Status" }, new[] { "Region", "Values" }, new[] { "Amount|Retail average|average|0.0",
                "OrderId|Retail order count|count|" }, "Channel", "Retail");
    }

    /// <summary>
    /// 验证指定透视的字段角色、筛选、数字格式、来源缓存及所有 OOXML 关系。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="targetSheet">透视目标工作表。</param>
    /// <param name="pivotName">透视名称。</param>
    /// <param name="sourceRange">来源单元格区域。</param>
    /// <param name="headers">来源字段顺序。</param>
    /// <param name="expectedRows">完整缓存记录。</param>
    /// <param name="rowFields">预期行轴字段。</param>
    /// <param name="columnFields">预期列轴字段。</param>
    /// <param name="valueFields">字段、标题、聚合和格式的预期组合。</param>
    /// <param name="filterField">报表筛选字段。</param>
    /// <param name="filterValue">报表筛选选项。</param>
    private static void AssertPivot(Dictionary<string, XDocument> package, string targetSheet,
        string pivotName, string sourceRange, IReadOnlyList<string> headers, IReadOnlyList<string> expectedRows,
        IReadOnlyList<string> rowFields, IReadOnlyList<string> columnFields,
        IReadOnlyList<string> valueFields, string filterField, string filterValue)
    {
        var worksheetPart = ResolveWorksheetPart(package, targetSheet);
        var pivotPart = ResolveRelatedPart(package, worksheetPart, "pivotTable");
        var pivot = package[pivotPart].Root;
        Assert.Equal(pivotName, (string)pivot.Attribute("name"));
        var location = (string)pivot.Element(SpreadsheetNs + "location").Attribute("ref");
        var expectedTarget = targetSheet == "RegionSummary" ? "C4" : "B5";
        Assert.StartsWith(expectedTarget, location);

        var cacheId = (string)pivot.Attribute("cacheId");
        var cachePart = ResolveCachePart(package, cacheId);
        var cache = package[cachePart].Root;
        var cacheFields = cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        Assert.Equal(headers, cacheFields.Select(field => (string)field.Attribute("name")));
        var source = cache.Element(SpreadsheetNs + "cacheSource")
            .Element(SpreadsheetNs + "worksheetSource");
        if (source.Attribute("sheet") != null)
        {
            Assert.Equal("Data", (string)source.Attribute("sheet"));
            Assert.Equal(sourceRange, (string)source.Attribute("ref"));
        }
        else
        {
            Assert.Equal("CompositionData", (string)source.Attribute("name"));
            var dataPart = ResolveWorksheetPart(package, "Data");
            var tablePart = ResolveRelatedPart(package, dataPart, "table");
            var sourceTable = package[tablePart].Root;
            Assert.Equal("CompositionData", (string)sourceTable.Attribute("displayName"));
            Assert.Equal(sourceRange, (string)sourceTable.Attribute("ref"));
        }

        var recordsPart = ResolveRelatedPart(package, cachePart, "pivotCacheRecords");
        var records = package[recordsPart].Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(expectedRows.Count, records.Length);
        Assert.All(records, row => Assert.Equal(headers.Count, row.Elements().Count()));
        Assert.Equal(expectedRows, records.Select(row => string.Join("|", Enumerable.Range(0, headers.Count)
            .Select(index => CacheValue(cacheFields[index], row.Elements().ElementAt(index))))));

        var pivotFields = pivot.Element(SpreadsheetNs + "pivotFields").Elements(SpreadsheetNs + "pivotField")
            .ToArray();
        Assert.Equal(rowFields, pivot.Element(SpreadsheetNs + "rowFields")
            .Elements(SpreadsheetNs + "field").Select(field => ResolvePivotFieldName(field, headers)));
        Assert.Equal(columnFields, pivot.Element(SpreadsheetNs + "colFields")
            .Elements(SpreadsheetNs + "field").Select(field => ResolvePivotFieldName(field, headers)));

        var actualValues = pivot.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField").Select(field =>
            {
                var index = int.Parse((string)field.Attribute("fld"));
                var format = ResolveNumberFormat(package, (string)field.Attribute("numFmtId"));
                return string.Join("|", headers[index], (string)field.Attribute("name"),
                    (string)field.Attribute("subtotal") ?? "sum", format);
            }).ToArray();
        Assert.Equal(valueFields, actualValues);

        var page = Assert.Single(pivot.Element(SpreadsheetNs + "pageFields")
            .Elements(SpreadsheetNs + "pageField"));
        var filterIndex = int.Parse((string)page.Attribute("fld"));
        Assert.Equal(filterField, headers[filterIndex]);
        var fieldItems = pivotFields[filterIndex].Element(SpreadsheetNs + "items")
            .Elements(SpreadsheetNs + "item").ToArray();
        var selected = fieldItems[int.Parse((string)page.Attribute("item"))];
        var sharedIndex = int.Parse((string)selected.Attribute("x"));
        Assert.Equal(filterValue, (string)cacheFields[filterIndex]
            .Element(SpreadsheetNs + "sharedItems").Elements().ElementAt(sharedIndex).Attribute("v"));

        Assert.StartsWith("xl/pivotTables/", pivotPart);
    }

    /// <summary>
    /// 从工作簿缓存关系解析透视表引用的缓存定义部件。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="cacheId">透视表缓存编号。</param>
    /// <returns>缓存定义部件路径。</returns>
    private static string ResolveCachePart(Dictionary<string, XDocument> package, string cacheId)
    {
        var cache = package["xl/workbook.xml"].Root.Element(SpreadsheetNs + "pivotCaches")
            .Elements(SpreadsheetNs + "pivotCache")
            .Single(item => (string)item.Attribute("cacheId") == cacheId);
        var relationshipId = (string)cache.Attribute(OfficeRelationshipNs + "id");
        var relationship = package["xl/_rels/workbook.xml.rels"].Root
            .Elements(PackageRelationshipNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == relationshipId);
        var target = ResolvePartPath("xl/", (string)relationship.Attribute("Target"));
        Assert.Contains(target, package.Keys);
        return target;
    }

    /// <summary>
    /// 按工作簿关系解析工作表的实际包路径。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <returns>工作表部件路径。</returns>
    private static string ResolveWorksheetPart(Dictionary<string, XDocument> package, string sheetName)
    {
        var sheet = package["xl/workbook.xml"].Root.Element(SpreadsheetNs + "sheets")
            .Elements(SpreadsheetNs + "sheet")
            .Single(item => (string)item.Attribute("name") == sheetName);
        var relationshipId = (string)sheet.Attribute(OfficeRelationshipNs + "id");
        var relationship = package["xl/_rels/workbook.xml.rels"].Root
            .Elements(PackageRelationshipNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == relationshipId);
        return ResolvePartPath("xl/", (string)relationship.Attribute("Target"));
    }

    /// <summary>
    /// 解析源部件到目标部件的指定类型关系。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="sourcePart">关系源部件路径。</param>
    /// <param name="relationshipSuffix">关系类型后缀。</param>
    /// <returns>关系目标部件路径。</returns>
    private static string ResolveRelatedPart(Dictionary<string, XDocument> package,
        string sourcePart, string relationshipSuffix)
    {
        var slash = sourcePart.LastIndexOf('/');
        var directory = sourcePart.Substring(0, slash + 1);
        var relationshipPart = directory + "_rels/" + sourcePart.Substring(slash + 1) + ".rels";
        var relationship = Assert.Single(package[relationshipPart].Root
            .Elements(PackageRelationshipNs + "Relationship")
            .Where(item => ((string)item.Attribute("Type"))?.EndsWith("/" + relationshipSuffix,
                StringComparison.Ordinal) == true));
        var target = ResolvePartPath(directory, (string)relationship.Attribute("Target"));
        Assert.Contains(target, package.Keys);
        return target;
    }

    /// <summary>
    /// 解析工作簿包关系中的相对目标路径。
    /// </summary>
    /// <param name="sourceDirectory">关系源部件目录。</param>
    /// <param name="target">关系目标路径。</param>
    /// <returns>规范化后的包内路径。</returns>
    private static string ResolvePartPath(string sourceDirectory, string target)
    {
        var segments = (target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : sourceDirectory + target)
            .Split('/');
        var resolved = new List<string>();
        foreach (var segment in segments)
        {
            if (segment == "..")
                resolved.RemoveAt(resolved.Count - 1);
            else if (segment != "." && segment.Length > 0)
                resolved.Add(segment);
        }
        return string.Join("/", resolved);
    }

    /// <summary>
    /// 解析按索引引用的缓存字符串或直接存储值。
    /// </summary>
    /// <param name="cacheField">对应来源字段的缓存定义。</param>
    /// <param name="recordCell">缓存记录中的单元格节点。</param>
    /// <returns>与来源工作表相同的文本表示。</returns>
    private static string CacheValue(XElement cacheField, XElement recordCell)
    {
        if (recordCell.Name == SpreadsheetNs + "m")
            return string.Empty;
        if (recordCell.Name == SpreadsheetNs + "x")
        {
            var index = int.Parse((string)recordCell.Attribute("v"));
            return (string)cacheField.Element(SpreadsheetNs + "sharedItems")
                .Elements().ElementAt(index).Attribute("v");
        }
        return (string)recordCell.Attribute("v") ?? recordCell.Value;
    }

    /// <summary>
    /// 将透视数字格式 ID 解析为实际格式代码。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="formatId">透视值字段使用的数字格式 ID。</param>
    /// <returns>格式代码；默认格式返回空字符串。</returns>
    private static string ResolveNumberFormat(Dictionary<string, XDocument> package, string formatId)
    {
        if (formatId == null)
            return string.Empty;
        if (formatId == "0")
            return string.Empty;
        if (formatId == "2")
            return "0.00";
        var formats = package["xl/styles.xml"].Root.Element(SpreadsheetNs + "numFmts")?
            .Elements(SpreadsheetNs + "numFmt") ?? Enumerable.Empty<XElement>();
        return (string)formats.Single(item => (string)item.Attribute("numFmtId") == formatId)
            .Attribute("formatCode");
    }

    /// <summary>
    /// 将透视轴字段索引解析为来源标题或合成值字段。
    /// </summary>
    /// <param name="field">透视行轴或列轴定义节点。</param>
    /// <param name="headers">来源字段标题顺序。</param>
    /// <returns>对应来源字段名，或多个度量使用的 Values 标记。</returns>
    private static string ResolvePivotFieldName(XElement field, IReadOnlyList<string> headers)
    {
        var index = int.Parse((string)field.Attribute("x"));
        return index == -2 ? "Values" : headers[index];
    }

    /// <summary>
    /// 读取 XLSX 中所有 XML 部件并保留完整路径。
    /// </summary>
    /// <param name="bytes">完整 XLSX 文件字节。</param>
    /// <returns>包内路径到 XML 文档的映射。</returns>
    private static Dictionary<string, XDocument> ReadPackage(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var part = entry.Open();
                return XDocument.Load(part);
            }, StringComparer.Ordinal);
    }

    /// <summary>
    /// 按透视表名称读取其定义 XML。
    /// </summary>
    /// <param name="package">完整 XLSX XML 部件。</param>
    /// <param name="name">透视表名称。</param>
    /// <returns>透视表 XML 文档。</returns>
    private static XDocument GetPivotByName(Dictionary<string, XDocument> package, string name) =>
        Assert.Single(package.Where(part => part.Key.StartsWith("xl/pivotTables/", StringComparison.Ordinal)
                && (string)part.Value.Root.Attribute("name") == name)
            .Select(part => part.Value));
}
