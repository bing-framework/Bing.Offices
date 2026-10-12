using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Tests.Exports.PivotTables.Business;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证透视来源组合动态列转换、工作表保护和表格配置。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotBusinessCompositionTest
{
    /// <summary>
    /// SpreadsheetML 工作簿结构使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Office Open XML 关系使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace OfficeRelationshipNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// Open Packaging Conventions 关系使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace PackageRelationshipNs =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// 验证公开同步和异步导出均保留转换后的完整透视来源。
    /// </summary>
    /// <param name="asynchronous">是否通过异步导出入口写入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_DynamicConverterWithProtectionAndTable_ShouldPreserveCompletePivotSource(
        bool asynchronous)
    {
        var request = CreateRequest();
        var exporter = new NpoiExcelExporter(new IExcelValueConverter[]
        {
            new NpoiPivotBusinessChannelConverter()
        });
        using var output = new MemoryStream();

        if (asynchronous)
            await exporter.ExportAsync(request, output);
        else
            exporter.Export(request, output);

        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        AssertComposition(output.ToArray());
    }

    /// <summary>
    /// 验证业务列映射、动态转换列、保护范围、表格和高级透视配置。
    /// </summary>
    /// <returns>由公开工作簿构建器生成的导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest()
    {
        var rows = new[]
        {
            new NpoiPivotBusinessRow
            {
                OrderNo = "SO-01", Amount = 10, Status = "Open",
                DynamicValues = new Dictionary<string, object> { ["channelCode"] = "R" }
            },
            new NpoiPivotBusinessRow
            {
                OrderNo = "SO-02", Amount = 20, Status = "Closed",
                DynamicValues = new Dictionary<string, object> { ["channelCode"] = "W" }
            },
            new NpoiPivotBusinessRow
            {
                OrderNo = "SO-03", Amount = 30, Status = "Open",
                DynamicValues = new Dictionary<string, object> { ["channelCode"] = "W" }
            }
        };
        var sourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 3
        };

        return ExcelExport.Workbook(builder => builder
            .AddSheet("Sales", rows, sheet => sheet
                .DynamicColumns(row => row.DynamicValues, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "channelCode",
                        Title = "Channel",
                        DataType = typeof(string),
                        ConverterName = "pivot-business-channel"
                    }
                })
                .Table(new ExcelTableDefinition
                {
                    Name = "SalesTable",
                    Range = sourceRange,
                    StyleName = "TableStyleMedium2"
                })
                .Protection(new ExcelSheetProtectionDefinition(new[]
                {
                    new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 3, StartColumn = 0, EndColumn = 3
                    }
                }, "sales-password")))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByChannel",
                SourceSheet = "Sales",
                SourceRange = sourceRange,
                TargetSheet = "Summary",
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Channel" }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" }
                },
                Filters = new[]
                {
                    new ExcelPivotFilterDefinition
                    {
                        Field = "Status", Values = new object[] { "Open" }
                    }
                }
            }));
    }

    /// <summary>
    /// 核对实际工作表单元格、缓存记录、透视角色和部件关系。
    /// </summary>
    /// <param name="bytes">由公开 NPOI 导出器生成的完整 XLSX 文件。</param>
    private static void AssertComposition(byte[] bytes)
    {
        using (var workbook = new XSSFWorkbook(new MemoryStream(bytes, writable: false)))
        {
            Assert.Equal(new[] { "Sales", "Summary" },
                Enumerable.Range(0, workbook.NumberOfSheets).Select(workbook.GetSheetName));
            var sales = workbook.GetSheet("Sales");
            var expectedHeaders = new[] { "OrderNo", "Amount", "Status", "Channel" };
            Assert.Equal(expectedHeaders, Enumerable.Range(0, expectedHeaders.Length)
                .Select(column => sales.GetRow(0).GetCell(column).StringCellValue));

            var expectedRows = ExpectedRows();
            for (var index = 0; index < expectedRows.Length; index++)
            {
                var row = sales.GetRow(index + 1);
                Assert.Equal(expectedRows[index].OrderNo, row.GetCell(0).StringCellValue);
                Assert.Equal(expectedRows[index].Amount, row.GetCell(1).NumericCellValue);
                Assert.Equal(expectedRows[index].Status, row.GetCell(2).StringCellValue);
                Assert.Equal(expectedRows[index].DynamicValues["channelCode"], row.GetCell(3).StringCellValue);
            }

            Assert.True(sales.Protect);
            Assert.True(sales.GetRow(0).GetCell(0).CellStyle.IsLocked);
            Assert.False(sales.GetRow(1).GetCell(0).CellStyle.IsLocked);
            Assert.False(sales.GetRow(1).GetCell(3).CellStyle.IsLocked);
            Assert.False(sales.GetRow(3).GetCell(3).CellStyle.IsLocked);

            var table = Assert.Single(((XSSFSheet)sales).GetTables());
            Assert.Equal("SalesTable", table.Name);
            Assert.Equal("A1:D4", table.CellReferences.FormatAsString());
            Assert.Equal("TableStyleMedium2", table.GetCTTable().tableStyleInfo.name);

            var pivots = Assert.Single(((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables());
            Assert.Equal("SalesByChannel", pivots.GetCTPivotTableDefinition().name);
        }

        AssertCacheAndPivotParts(bytes);
    }

    /// <summary>
    /// 核对缓存中的所有来源字段和记录，包括筛选未选中的订单。
    /// </summary>
    /// <param name="bytes">完整 XLSX 文件。</param>
    private static void AssertCacheAndPivotParts(byte[] bytes)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var package = archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || entry.FullName.EndsWith(".rels", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, ReadXml);
        var workbookXml = package["xl/workbook.xml"];
        var workbookRels = package["xl/_rels/workbook.xml.rels"];
        var salesSheet = workbookXml.Root.Element(SpreadsheetNs + "sheets")
            .Elements(SpreadsheetNs + "sheet").Single(sheet => (string)sheet.Attribute("name") == "Sales");
        var salesSheetId = (string)salesSheet.Attribute(OfficeRelationshipNs + "id");
        var salesSheetTarget = workbookRels.Root.Elements(PackageRelationshipNs + "Relationship")
            .Single(relation => (string)relation.Attribute("Id") == salesSheetId)
            .Attribute("Target").Value;
        var salesSheetPart = ResolvePart("xl/", salesSheetTarget);
        var salesRelationshipsPart = RelatedPart(salesSheetPart);
        var salesRelationships = package[salesRelationshipsPart].Root
            .Elements(PackageRelationshipNs + "Relationship").ToArray();
        var tableRelationship = Assert.Single(salesRelationships.Where(relation =>
            ((string)relation.Attribute("Type"))?.EndsWith("/table", StringComparison.Ordinal) == true));
        var tablePart = ResolvePart(DirectoryOf(salesSheetPart), (string)tableRelationship.Attribute("Target"));
        Assert.Equal("SalesTable", (string)package[tablePart].Root.Attribute("displayName"));
        Assert.Equal("A1:D4", (string)package[tablePart].Root.Attribute("ref"));

        var summarySheet = workbookXml.Root.Element(SpreadsheetNs + "sheets")
            .Elements(SpreadsheetNs + "sheet").Single(sheet => (string)sheet.Attribute("name") == "Summary");
        var summarySheetId = (string)summarySheet.Attribute(OfficeRelationshipNs + "id");
        var summaryTarget = workbookRels.Root.Elements(PackageRelationshipNs + "Relationship")
            .Single(relation => (string)relation.Attribute("Id") == summarySheetId)
            .Attribute("Target").Value;
        var summaryPart = ResolvePart("xl/", summaryTarget);
        var summaryRelsPart = RelatedPart(summaryPart);
        var summaryRelationships = package[summaryRelsPart].Root
            .Elements(PackageRelationshipNs + "Relationship").ToArray();
        var pivotRelation = Assert.Single(summaryRelationships.Where(relation =>
            ((string)relation.Attribute("Type"))?.EndsWith("/pivotTable", StringComparison.Ordinal) == true));
        var pivotPart = ResolvePart(DirectoryOf(summaryPart), (string)pivotRelation.Attribute("Target"));
        var pivot = package[pivotPart].Root;

        Assert.Equal("SalesByChannel", (string)pivot.Attribute("name"));
        var pivotRowField = Assert.Single(pivot.Element(SpreadsheetNs + "rowFields").Elements());
        Assert.Equal(3, (int)pivotRowField.Attribute("x"));
        var dataField = Assert.Single(pivot.Element(SpreadsheetNs + "dataFields").Elements());
        Assert.Equal(1, (int)dataField.Attribute("fld"));
        Assert.Equal("sum", (string)dataField.Attribute("subtotal") ?? "sum");
        var pageField = Assert.Single(pivot.Element(SpreadsheetNs + "pageFields").Elements());
        Assert.Equal(2, (int)pageField.Attribute("fld"));
        var statusItems = pivot.Element(SpreadsheetNs + "pivotFields").Elements(SpreadsheetNs + "pivotField")
            .ElementAt(2).Element(SpreadsheetNs + "items").Elements(SpreadsheetNs + "item").ToArray();
        var selectedStatusItem = statusItems[(int)pageField.Attribute("item")];
        var statusSharedIndex = (int)selectedStatusItem.Attribute("x");

        var cacheId = (string)pivot.Attribute("cacheId");
        var workbookCache = workbookXml.Root.Element(SpreadsheetNs + "pivotCaches")
            .Elements(SpreadsheetNs + "pivotCache").Single(cache => (string)cache.Attribute("cacheId") == cacheId);
        var cacheRelationshipId = (string)workbookCache.Attribute(OfficeRelationshipNs + "id");
        var cacheTarget = workbookRels.Root.Elements(PackageRelationshipNs + "Relationship")
            .Single(relation => (string)relation.Attribute("Id") == cacheRelationshipId)
            .Attribute("Target").Value;
        var cachePart = ResolvePart("xl/", cacheTarget);
        var cacheRelsPart = RelatedPart(cachePart);
        var cacheRelationships = package[cacheRelsPart].Root
            .Elements(PackageRelationshipNs + "Relationship").ToArray();
        var recordsRelation = Assert.Single(cacheRelationships.Where(relation =>
            ((string)relation.Attribute("Type"))?.EndsWith("/pivotCacheRecords", StringComparison.Ordinal) == true));
        var recordsPart = ResolvePart(DirectoryOf(cachePart), (string)recordsRelation.Attribute("Target"));

        var cache = package[cachePart].Root;
        var fields = cache.Element(SpreadsheetNs + "cacheFields").Elements(SpreadsheetNs + "cacheField").ToArray();
        Assert.Equal(new[] { "OrderNo", "Amount", "Status", "Channel" },
            fields.Select(field => (string)field.Attribute("name")));
        var source = cache.Element(SpreadsheetNs + "cacheSource").Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Sales", (string)source.Attribute("sheet"));
        Assert.Equal("A1:D4", (string)source.Attribute("ref"));
        Assert.Equal("Open", (string)fields[2].Element(SpreadsheetNs + "sharedItems")
            .Elements().ElementAt(statusSharedIndex).Attribute("v"));
        Assert.Equal(3, (int)cache.Attribute("recordCount"));

        var records = package[recordsPart].Root.Elements(SpreadsheetNs + "r").ToArray();
        var expectedRows = ExpectedRows();
        Assert.Equal(expectedRows.Length, records.Length);
        var expectedMatrix = expectedRows.Select(row => new object[]
        {
            row.OrderNo, row.Amount, row.Status, row.DynamicValues["channelCode"]
        }).ToArray();
        for (var rowIndex = 0; rowIndex < records.Length; rowIndex++)
        {
            var indexes = records[rowIndex].Elements(SpreadsheetNs + "x")
                .Select(item => (int)item.Attribute("v")).ToArray();
            Assert.Equal(fields.Length, indexes.Length);
            for (var column = 0; column < fields.Length; column++)
            {
                var item = fields[column].Element(SpreadsheetNs + "sharedItems").Elements().ElementAt(indexes[column]);
                Assert.Equal(expectedMatrix[rowIndex][column], ReadSharedItem(item));
            }
        }

        var contentTypes = package["[Content_Types].xml"].Root;
        Assert.Contains(contentTypes.Elements().Where(element => element.Name.LocalName == "Override"), element =>
            (string)element.Attribute("PartName") == "/" + pivotPart);
        Assert.Contains(contentTypes.Elements().Where(element => element.Name.LocalName == "Override"), element =>
            (string)element.Attribute("PartName") == "/" + cachePart);
        Assert.Contains(contentTypes.Elements().Where(element => element.Name.LocalName == "Override"), element =>
            (string)element.Attribute("PartName") == "/" + recordsPart);
        Assert.Contains(contentTypes.Elements().Where(element => element.Name.LocalName == "Override"), element =>
            (string)element.Attribute("PartName") == "/" + tablePart);
        Assert.Contains(pivotPart, package.Keys);
        Assert.Contains(cachePart, package.Keys);
        Assert.Contains(recordsPart, package.Keys);
        Assert.Contains(tablePart, package.Keys);
    }

    /// <summary>
    /// 读取一个 XLSX 包部件并解析其 XML。
    /// </summary>
    /// <param name="entry">待读取的压缩包部件。</param>
    /// <returns>已加载的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 将工作表或缓存部件的相对关系目标转换为包路径。
    /// </summary>
    /// <param name="baseDirectory">关系源部件所在目录。</param>
    /// <param name="target">关系目标路径。</param>
    /// <returns>规范化后的包内路径。</returns>
    private static string ResolvePart(string baseDirectory, string target)
    {
        var combined = target.StartsWith("/", StringComparison.Ordinal)
            ? target.Substring(1)
            : baseDirectory + target;
        return string.Join("/", combined.Split('/').Aggregate(new List<string>(), (parts, part) =>
        {
            if (part == "..")
                parts.RemoveAt(parts.Count - 1);
            else if (part.Length != 0 && part != ".")
                parts.Add(part);
            return parts;
        }));
    }

    /// <summary>
    /// 返回部件所在目录的包路径前缀。
    /// </summary>
    /// <param name="part">源部件路径。</param>
    /// <returns>以斜线结尾的目录前缀。</returns>
    private static string DirectoryOf(string part) => part.Substring(0, part.LastIndexOf('/') + 1);

    /// <summary>
    /// 生成指定源部件的关系部件路径。
    /// </summary>
    /// <param name="part">源部件路径。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelatedPart(string part)
    {
        var slash = part.LastIndexOf('/');
        return part.Substring(0, slash + 1) + "_rels/" + part.Substring(slash + 1) + ".rels";
    }

    /// <summary>
    /// 解码缓存字段中的共享项值。
    /// </summary>
    /// <param name="item">共享项 XML 元素。</param>
    /// <returns>按 SpreadsheetML 类型解码的业务值。</returns>
    private static object ReadSharedItem(XElement item)
    {
        if (item.Name == SpreadsheetNs + "s")
            return (string)item.Attribute("v");
        if (item.Name == SpreadsheetNs + "n")
            return (double)item.Attribute("v");
        throw new InvalidDataException($"透视缓存包含未预期的共享项类型 {item.Name}。");
    }

    /// <summary>
    /// 创建预期的转换后来源行。
    /// </summary>
    /// <returns>按原始顺序排列的完整透视来源。</returns>
    private static NpoiPivotBusinessRow[] ExpectedRows() => new[]
    {
        new NpoiPivotBusinessRow { OrderNo = "SO-01", Amount = 10, Status = "Open",
            DynamicValues = new Dictionary<string, object> { ["channelCode"] = "Retail" } },
        new NpoiPivotBusinessRow { OrderNo = "SO-02", Amount = 20, Status = "Closed",
            DynamicValues = new Dictionary<string, object> { ["channelCode"] = "Wholesale" } },
        new NpoiPivotBusinessRow { OrderNo = "SO-03", Amount = 30, Status = "Open",
            DynamicValues = new Dictionary<string, object> { ["channelCode"] = "Wholesale" } }
    };
}
