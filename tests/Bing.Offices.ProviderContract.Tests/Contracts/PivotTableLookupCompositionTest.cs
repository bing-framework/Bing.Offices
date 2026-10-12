using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证透视目标工作表与 LookupList、级联字典辅助工作表的组合命名。
/// </summary>
public sealed class PivotTableLookupCompositionTest
{
    /// <summary>
    /// 验证透视目标以不同大小写占用辅助名称时，同步和异步导出均顺延字典工作表。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_PivotTargets_ShouldReserveLookupAndCascadeSheets(bool asynchronous)
    {
        using var output = new MemoryStream();
        await ExportAsync(CreateRequest(), output, asynchronous);

        Assert.True(output.CanWrite);
        using var workbook = OpenWorkbook(output);
        AssertWorkbookContents(workbook, template: false);
        AssertPivotParts(output.ToArray(), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["__BingLookup_001"] = "RegionTotalByLookup",
            ["__bingcascade_001"] = "RegionTotalByCascade"
        });
    }

    /// <summary>
    /// 验证模板已有 001 且透视目标占用 002 时保留模板并将辅助工作表顺延。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_TemplateAndPivotTargets_ShouldPreserveTemplateAndAdvanceNames(bool asynchronous)
    {
        using var template = CreateTemplate();
        var request = CreateRequest(template);
        using var output = new MemoryStream();
        await ExportAsync(request, output, asynchronous);

        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
        using var workbook = OpenWorkbook(output);
        AssertWorkbookContents(workbook, template: true);
        AssertPivotParts(output.ToArray(), new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["__BingLookup_002"] = "RegionTotalByLookup",
            ["__BingCascade_002"] = "RegionTotalByCascade"
        });
    }

    /// <summary>
    /// 创建同时包含普通字典、级联字典和共源透视表的工作簿请求。
    /// </summary>
    /// <param name="template">可选模板流。</param>
    /// <returns>已构建的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template = null) =>
        ExcelExport.Workbook(builder =>
        {
            if (template != null)
                builder.UseTemplate(template, leaveOpen: true);
            builder
                .LookupList("OrderStatus", new[] { "Open", "Closed" })
                .LookupList("Priority", new[] { "High", "Low" })
                .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
                {
                    ["Hardware"] = new[] { "Chair", "Desk" }
                })
                .CascadeList("Divisions", new Dictionary<string, IReadOnlyList<string>>
                {
                    ["North"] = new[] { "Depot-A", "Depot-B" }
                })
                .AddSheet("Sales", new[]
                {
                    new SalesRow { Region = "North", Amount = 10 },
                    new SalesRow { Region = "South", Amount = 20 },
                    new SalesRow { Region = "North", Amount = 30 }
                })
                .AddSheet("Entry", new[]
                {
                    new EntryRow
                    {
                        Status = "Open", Priority = "High", Category = "Hardware", Product = "Chair",
                        Division = "North", Depot = "Depot-A"
                    },
                    new EntryRow
                    {
                        Status = "Closed", Priority = "Low", Category = "Hardware", Product = "Desk",
                        Division = "North", Depot = "Depot-B"
                    }
                }, sheet => sheet
                    .DataValidation(ListValidation("OrderStatus", 0))
                    .DataValidation(ListValidation("Priority", 1))
                    .DataValidation(CascadeValidation("Categories", 2, 3))
                    .DataValidation(CascadeValidation("Divisions", 4, 5)))
                .PivotTable(Pivot("RegionTotalByLookup", template == null
                    ? "__BingLookup_001" : "__BingLookup_002"))
                .PivotTable(Pivot("RegionTotalByCascade", template == null
                    ? "__bingcascade_001" : "__BingCascade_002"));
        });

    /// <summary>
    /// 创建区域求和透视表定义。
    /// </summary>
    /// <param name="name">透视表名称。</param>
    /// <param name="targetSheet">透视表目标工作表。</param>
    /// <returns>指向 Sales 明细的透视表定义。</returns>
    private static ExcelPivotTableDefinition Pivot(string name, string targetSheet) => new()
    {
        Name = name,
        SourceSheet = "Sales",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
        },
        TargetSheet = targetSheet,
        TargetRow = 1,
        TargetColumn = 2,
        RowField = "Region",
        ValueField = "Amount"
    };

    /// <summary>
    /// 创建工作簿名称列表校验定义。
    /// </summary>
    /// <param name="sourceName">字典名称。</param>
    /// <param name="column">目标列索引。</param>
    /// <returns>覆盖 Entry 两条明细的列表校验定义。</returns>
    private static ExcelDataValidationDefinition ListValidation(string sourceName, int column) => new()
    {
        Range = new ExcelRangeDefinition
        {
            StartRow = 1, EndRow = 2, StartColumn = column, EndColumn = column
        },
        Type = ExcelDataValidationType.List,
        ListSourceName = sourceName
    };

    /// <summary>
    /// 创建按同一行上级值选择子列表的校验定义。
    /// </summary>
    /// <param name="sourceName">级联字典名称。</param>
    /// <param name="parentColumn">上级值所在列。</param>
    /// <param name="childColumn">子列表校验所在列。</param>
    /// <returns>覆盖 Entry 两条明细的级联校验定义。</returns>
    private static ExcelDataValidationDefinition CascadeValidation(string sourceName,
        int parentColumn, int childColumn) => new()
    {
        Range = new ExcelRangeDefinition
        {
            StartRow = 1, EndRow = 2, StartColumn = childColumn, EndColumn = childColumn
        },
        Type = ExcelDataValidationType.List,
        CascadeSourceName = sourceName,
        ParentColumnIndex = parentColumn
    };

    /// <summary>
    /// 按同步或异步入口导出工作簿。
    /// </summary>
    /// <param name="request">待导出的工作簿请求。</param>
    /// <param name="output">由测试持有的目标流。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    private static async Task ExportAsync(ExcelWorkbookExportRequest request, Stream output, bool asynchronous)
    {
        var exporter = new ClosedXmlExcelExporter();
        if (asynchronous)
            await exporter.ExportAsync(request, output);
        else
            exporter.Export(request, output);
    }

    /// <summary>
    /// 从目标流打开导出工作簿。
    /// </summary>
    /// <param name="output">包含 XLSX 数据的目标流。</param>
    /// <returns>已加载的工作簿。</returns>
    private static XLWorkbook OpenWorkbook(MemoryStream output) =>
        new(new MemoryStream(output.ToArray(), writable: false));

    /// <summary>
    /// 验证业务数据、辅助工作表、名称、校验公式及模板内容。
    /// </summary>
    /// <param name="workbook">导出后的工作簿。</param>
    /// <param name="template">是否从模板导出。</param>
    private static void AssertWorkbookContents(XLWorkbook workbook, bool template)
    {
        var lookupStatusSheet = template ? "__BingLookup_003" : "__BingLookup_002";
        var lookupPrioritySheet = template ? "__BingLookup_004" : "__BingLookup_003";
        var cascadeCategorySheet = template ? "__BingCascade_003" : "__BingCascade_002";
        var cascadeDivisionSheet = template ? "__BingCascade_004" : "__BingCascade_003";
        var expectedSheets = new[] { "Sales", "Entry", lookupStatusSheet, lookupPrioritySheet,
            cascadeCategorySheet, cascadeDivisionSheet,
            template ? "Notes" : null,
            template ? "__BingLookup_001" : null,
            template ? "__BingCascade_001" : null,
            template ? "__BingLookup_002" : "__BingLookup_001",
            template ? "__BingCascade_002" : "__bingcascade_001" }
            .Where(name => name != null).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray();
        Assert.Equal(expectedSheets, workbook.Worksheets.Select(sheet => sheet.Name)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase));
        var sheetOrder = workbook.Worksheets.Select(sheet => sheet.Name).ToArray();
        Assert.True(Array.IndexOf(sheetOrder, "Sales") < Array.IndexOf(sheetOrder, "Entry"));
        if (template)
            Assert.Equal(new[] { "Sales", "Entry", "Notes", "__BingLookup_001", "__BingCascade_001" },
                sheetOrder.Take(5));

        var sales = workbook.Worksheet("Sales");
        Assert.Equal(new[] { "Region", "Amount", "North", "10", "South", "20", "North", "30" },
            sales.Range("A1:B4").Cells().Select(cell => cell.GetString()));
        var entry = workbook.Worksheet("Entry");
        Assert.Equal(new[]
        {
            "Status", "Priority", "Category", "Product", "Division", "Depot",
            "Open", "High", "Hardware", "Chair", "North", "Depot-A",
            "Closed", "Low", "Hardware", "Desk", "North", "Depot-B"
        }, entry.Range("A1:F3").Cells().Select(cell => cell.GetString()));

        AssertLookupSheet(workbook, lookupStatusSheet, "Open", "Closed");
        AssertLookupSheet(workbook, lookupPrioritySheet, "High", "Low");
        AssertCascadeSheet(workbook, cascadeCategorySheet, "Hardware", "Categories",
            "__BingCascadeValues_002", "Chair", "Desk");
        AssertCascadeSheet(workbook, cascadeDivisionSheet, "North", "Divisions",
            "__BingCascadeValues_004", "Depot-A", "Depot-B");

        var names = workbook.DefinedNames.ValidNamedRanges().ToDictionary(item => item.Name,
            StringComparer.OrdinalIgnoreCase);
        AssertDefinedRange(names, "OrderStatus", lookupStatusSheet, 1, 1, 2, 1);
        AssertDefinedRange(names, "Priority", lookupPrioritySheet, 1, 1, 2, 1);
        AssertDefinedRange(names, "Categories", cascadeCategorySheet, 1, 1, 1, 1);
        AssertDefinedRange(names, "Divisions", cascadeDivisionSheet, 1, 1, 1, 1);
        AssertDefinedRange(names, "__BingCascadeMap_001", cascadeCategorySheet, 1, 1, 1, 2);
        AssertDefinedRange(names, "__BingCascadeValues_002", cascadeCategorySheet, 1, 3, 2, 3);
        AssertDefinedRange(names, "__BingCascadeMap_003", cascadeDivisionSheet, 1, 1, 1, 2);
        AssertDefinedRange(names, "__BingCascadeValues_004", cascadeDivisionSheet, 1, 3, 2, 3);

        var validations = entry.DataValidations.ToArray();
        Assert.Equal(4, validations.Length);
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') == "OrderStatus");
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') == "Priority");
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') ==
            "INDIRECT(VLOOKUP($C2,__BingCascadeMap_001,2,FALSE))");
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') ==
            "INDIRECT(VLOOKUP($E2,__BingCascadeMap_003,2,FALSE))");

        if (!template)
            return;
        AssertTemplateSheet(workbook, "Notes", "retained note");
        AssertTemplateSheet(workbook, "__BingLookup_001", "old lookup");
        AssertTemplateSheet(workbook, "__BingCascade_001", "old cascade");
    }

    /// <summary>
    /// 验证普通名称列表的所有字典文本与隐藏状态。
    /// </summary>
    /// <param name="workbook">导出后的工作簿。</param>
    /// <param name="sheetName">辅助工作表名称。</param>
    /// <param name="first">首个字典值。</param>
    /// <param name="second">第二个字典值。</param>
    private static void AssertLookupSheet(XLWorkbook workbook, string sheetName, string first, string second)
    {
        var sheet = workbook.Worksheet(sheetName);
        Assert.Equal(XLWorksheetVisibility.Hidden, sheet.Visibility);
        Assert.Equal(new[] { first, second }, new[] { sheet.Cell("A1").GetString(), sheet.Cell("A2").GetString() });
        Assert.True(sheet.Cell("A3").IsEmpty());
    }

    /// <summary>
    /// 验证级联辅助工作表中的上级、子名称及全部下级文本。
    /// </summary>
    /// <param name="workbook">导出后的工作簿。</param>
    /// <param name="sheetName">辅助工作表名称。</param>
    /// <param name="parent">上级文本。</param>
    /// <param name="parentName">上级名称。</param>
    /// <param name="childName">生成的下级名称。</param>
    /// <param name="firstChild">首个下级文本。</param>
    /// <param name="secondChild">第二个下级文本。</param>
    private static void AssertCascadeSheet(XLWorkbook workbook, string sheetName, string parent,
        string parentName, string childName, string firstChild, string secondChild)
    {
        var sheet = workbook.Worksheet(sheetName);
        Assert.Equal(XLWorksheetVisibility.Hidden, sheet.Visibility);
        Assert.Equal(parent, sheet.Cell("A1").GetString());
        Assert.Equal(childName, sheet.Cell("B1").GetString());
        Assert.Equal(new[] { firstChild, secondChild },
            new[] { sheet.Cell("C1").GetString(), sheet.Cell("C2").GetString() });
        Assert.True(sheet.Cell("A2").IsEmpty());
        Assert.True(sheet.Cell("B2").IsEmpty());
        Assert.True(sheet.Cell("C3").IsEmpty());
        Assert.Equal(parentName, Assert.Single(workbook.DefinedNames.ValidNamedRanges()
            .Where(item => item.Name == parentName)).Name);
    }

    /// <summary>
    /// 验证名称区域精确指向指定工作表及单元格范围。
    /// </summary>
    /// <param name="names">工作簿有效名称集合。</param>
    /// <param name="name">名称标识。</param>
    /// <param name="sheetName">目标工作表。</param>
    /// <param name="firstRow">首行号。</param>
    /// <param name="firstColumn">首列号。</param>
    /// <param name="lastRow">末行号。</param>
    /// <param name="lastColumn">末列号。</param>
    private static void AssertDefinedRange(IDictionary<string, IXLDefinedName> names, string name,
        string sheetName, int firstRow, int firstColumn, int lastRow, int lastColumn)
    {
        var range = Assert.Single(names[name].Ranges);
        Assert.Equal(sheetName, range.Worksheet.Name);
        Assert.Equal((firstRow, firstColumn, lastRow, lastColumn),
            (range.RangeAddress.FirstAddress.RowNumber, range.RangeAddress.FirstAddress.ColumnNumber,
                range.RangeAddress.LastAddress.RowNumber, range.RangeAddress.LastAddress.ColumnNumber));
    }

    /// <summary>
    /// 验证模板中的无关文本及格式仍然存在。
    /// </summary>
    /// <param name="workbook">导出后的工作簿。</param>
    /// <param name="sheetName">模板工作表名称。</param>
    /// <param name="value">模板原有文本。</param>
    private static void AssertTemplateSheet(XLWorkbook workbook, string sheetName, string value)
    {
        var cell = workbook.Worksheet(sheetName).Cell("Z9");
        Assert.Equal(value, cell.GetString());
        Assert.True(cell.Style.Font.Bold);
        Assert.Equal(XLColor.LightYellow.Color, cell.Style.Fill.BackgroundColor.Color);
    }

    /// <summary>
    /// 验证透视表、缓存、记录及其 OOXML 关系均指向实际工作表。
    /// </summary>
    /// <param name="bytes">导出的 XLSX 字节。</param>
    /// <param name="targets">目标工作表与透视表名称对应关系。</param>
    private static void AssertPivotParts(byte[] bytes, IReadOnlyDictionary<string, string> targets)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var workbookXml = ReadXml(archive, "xl/workbook.xml");
        var workbookRelationships = ReadRelationships(archive, "xl/workbook.xml");
        var cacheTargets = workbookXml.Root.Element(SpreadsheetNs + "pivotCaches")
            .Elements(SpreadsheetNs + "pivotCache").ToDictionary(item => (string)item.Attribute("cacheId"), item =>
            {
                var relationship = FindRelationship(workbookRelationships,
                    (string)item.Attribute(OfficeRelationshipNs + "id"));
                Assert.EndsWith("/pivotCacheDefinition", (string)relationship.Attribute("Type"));
                return ResolvePartPath("xl/", (string)relationship.Attribute("Target"));
            });
        var pivotParts = archive.Entries.Where(entry => entry.FullName.Contains("pivotTables/", StringComparison.Ordinal)
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal))
            .Select(entry => entry.FullName).ToArray();
        Assert.Equal(targets.Count, pivotParts.Length);
        var sheets = workbookXml.Root.Element(SpreadsheetNs + "sheets").Elements(SpreadsheetNs + "sheet")
            .ToDictionary(item => (string)item.Attribute("name"), item =>
            {
                var relationship = FindRelationship(workbookRelationships,
                    (string)item.Attribute(OfficeRelationshipNs + "id"));
                Assert.EndsWith("/worksheet", (string)relationship.Attribute("Type"));
                return ResolvePartPath("xl/", (string)relationship.Attribute("Target"));
            }, StringComparer.OrdinalIgnoreCase);
        var cachePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pivotPath in pivotParts)
        {
            var pivot = ReadXml(archive, pivotPath).Root;
            var target = targets.Single(item => item.Value == (string)pivot.Attribute("name"));
            Assert.Equal("C2", ((string)pivot.Element(SpreadsheetNs + "location").Attribute("ref")).Split(':')[0]);
            var cacheId = (string)pivot.Attribute("cacheId");
            Assert.True(cacheTargets.TryGetValue(cacheId, out var cachePath),
                $"缺少 cacheId={cacheId} 对应的工作簿缓存关系。");
            cachePaths.Add(cachePath);
            var targetRelationship = SingleRelationshipTarget(archive, sheets[target.Key], "pivotTable");
            Assert.Equal(pivotPath, targetRelationship);
            AssertPivotCache(archive, cachePath);
        }

        Assert.NotEmpty(cachePaths);
        Assert.All(cachePaths, path => AssertPivotCache(archive, path));
    }

    /// <summary>
    /// 验证透视缓存包含完整源区域、字段、记录和缓存记录关系。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="cachePath">缓存定义部件路径。</param>
    private static void AssertPivotCache(ZipArchive archive, string cachePath)
    {
        var cache = ReadXml(archive, cachePath).Root;
        var source = cache.Element(SpreadsheetNs + "cacheSource").Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Sales", (string)source.Attribute("sheet"));
        Assert.Equal("A1:B4", (string)source.Attribute("ref"));
        Assert.Equal(new[] { "Region", "Amount" }, cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").Select(field => (string)field.Attribute("name")));

        var recordsPath = SingleRelationshipTarget(archive, cachePath, "pivotCacheRecords");
        var records = ReadXml(archive, recordsPath).Root;
        var rows = records.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, rows.Length);
        Assert.Equal(new[] { "10", "20", "30" }, rows
            .Select(row => (string)row.Elements().Last().Attribute("v")));
        var regionValues = cache.Element(SpreadsheetNs + "cacheFields").Elements().First()
            .Element(SpreadsheetNs + "sharedItems").Elements()
            .Select(item => (string)item.Attribute("v")).ToArray();
        Assert.Equal(new[] { "North", "South", "North" }, rows.Select(row => row.Elements().First())
            .Select(item => item.Name.LocalName == "x"
                ? regionValues[int.Parse((string)item.Attribute("v"))]
                : (string)item.Attribute("v")));
    }

    /// <summary>
    /// 读取工作簿关系部件并定位指定类型的唯一目标。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="sourcePart">关系源部件路径。</param>
    /// <param name="typeSuffix">关系类型后缀。</param>
    /// <returns>相对于压缩包根目录的目标部件路径。</returns>
    private static string SingleRelationshipTarget(ZipArchive archive, string sourcePart, string typeSuffix)
    {
        var relationships = ReadRelationships(archive, sourcePart)
            .Where(item => ((string)item.Attribute("Type"))?.EndsWith("/" + typeSuffix,
                StringComparison.Ordinal) == true).ToArray();
        var relationship = Assert.Single(relationships);
        var directory = sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1);
        return ResolvePartPath(directory, (string)relationship.Attribute("Target"));
    }

    /// <summary>
    /// 读取并返回指定 OOXML 部件的关系节点。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="sourcePart">关系源部件路径。</param>
    /// <returns>关系节点集合。</returns>
    private static IEnumerable<XElement> ReadRelationships(ZipArchive archive, string sourcePart)
    {
        XNamespace relationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var directory = sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1);
        var file = sourcePart.Substring(sourcePart.LastIndexOf('/') + 1);
        return ReadXml(archive, directory + "_rels/" + file + ".rels").Root
            .Elements(relationshipNs + "Relationship");
    }

    /// <summary>
    /// 根据标识查找 OOXML 关系节点。
    /// </summary>
    /// <param name="relationships">关系节点集合。</param>
    /// <param name="id">关系标识。</param>
    /// <returns>匹配的关系 XML 节点。</returns>
    private static XElement FindRelationship(IEnumerable<XElement> relationships, string id)
    {
        return relationships.Single(item => (string)item.Attribute("Id") == id);
    }

    /// <summary>
    /// 将 OOXML 关系目标解析为压缩包内路径。
    /// </summary>
    /// <param name="sourceDirectory">源部件目录。</param>
    /// <param name="target">关系目标。</param>
    /// <returns>规范化部件路径。</returns>
    private static string ResolvePartPath(string sourceDirectory, string target)
    {
        var segments = (target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : sourceDirectory + target).Split('/');
        var result = new List<string>();
        foreach (var segment in segments)
        {
            if (segment == "..")
                result.RemoveAt(result.Count - 1);
            else if (segment.Length > 0 && segment != ".")
                result.Add(segment);
        }
        return string.Join("/", result);
    }

    /// <summary>
    /// 读取压缩包中的 XML 部件。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="path">部件路径。</param>
    /// <returns>XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        Assert.NotNull(entry);
        using var input = entry.Open();
        return XDocument.Load(input);
    }

    /// <summary>
    /// SpreadsheetML 工作簿部件使用的命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// Office Open XML 部件关系引用使用的命名空间。
    /// </summary>
    private static readonly XNamespace OfficeRelationshipNs =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// 创建包含模板无关内容、样式及已占用辅助名称的工作簿流。
    /// </summary>
    /// <returns>位置重置到开头的 XLSX 模板流。</returns>
    private static MemoryStream CreateTemplate()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("Sales");
        workbook.Worksheets.Add("Entry");
        AddTemplateCell(workbook.Worksheets.Add("Notes"), "retained note");
        AddTemplateCell(workbook.Worksheets.Add("__BingLookup_001"), "old lookup");
        AddTemplateCell(workbook.Worksheets.Add("__BingCascade_001"), "old cascade");
        var lookup = workbook.Worksheet("__BingLookup_001");
        var cascade = workbook.Worksheet("__BingCascade_001");
        lookup.Visibility = XLWorksheetVisibility.Hidden;
        cascade.Visibility = XLWorksheetVisibility.Hidden;
        var template = new MemoryStream();
        workbook.SaveAs(template);
        template.Position = 0;
        return template;
    }

    /// <summary>
    /// 在模板工作表写入可验证保留状态的文本和样式。
    /// </summary>
    /// <param name="sheet">待写入的模板工作表。</param>
    /// <param name="value">模板原有文本。</param>
    private static void AddTemplateCell(IXLWorksheet sheet, string value)
    {
        var cell = sheet.Cell("Z9");
        cell.Value = value;
        cell.Style.Font.Bold = true;
        cell.Style.Fill.BackgroundColor = XLColor.LightYellow;
    }

    /// <summary>
    /// 用于生成共源透视表的销售明细。
    /// </summary>
    private sealed class SalesRow
    {
        /// <summary>获取或设置区域。</summary>
        public string Region { get; set; }

        /// <summary>获取或设置金额。</summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 用于验证普通和级联下拉的业务录入行。
    /// </summary>
    private sealed class EntryRow
    {
        /// <summary>获取或设置状态。</summary>
        public string Status { get; set; }

        /// <summary>获取或设置优先级。</summary>
        public string Priority { get; set; }

        /// <summary>获取或设置类别。</summary>
        public string Category { get; set; }

        /// <summary>获取或设置商品。</summary>
        public string Product { get; set; }

        /// <summary>获取或设置区域。</summary>
        public string Division { get; set; }

        /// <summary>获取或设置仓库。</summary>
        public string Depot { get; set; }
    }
}
