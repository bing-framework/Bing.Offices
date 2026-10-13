using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exports;
using ClosedXML.Excel;

namespace Bing.Offices.Tools.PivotClientProbe;

/// <summary>
/// 生成并检查高级透视客户端夹具。
/// </summary>
internal static class AdvancedPivotProbe
{
    /// <summary>
    /// 生成高级透视客户端夹具。
    /// </summary>
    /// <param name="output">新建 XLSX 的文件路径。</param>
    /// <param name="asynchronous">是否使用异步文件导出。</param>
    /// <remarks>
    /// 结构检查不代替客户端聚合、筛选和刷新验收；已有文件拒绝覆盖。
    /// </remarks>
    internal static async Task RunAsync(string output, bool asynchronous)
        => await RunAsync(output, asynchronous, new ClosedXmlExcelExporter()).ConfigureAwait(false);

    /// <summary>
    /// 生成高级透视客户端夹具。
    /// </summary>
    /// <param name="output">新建 XLSX 的文件路径。</param>
    /// <param name="asynchronous">是否使用异步文件导出。</param>
    /// <param name="exporter">当前候选包提供的公开导出服务。</param>
    /// <remarks>
    /// 结构检查不代替客户端聚合、筛选和刷新验收；已有文件拒绝覆盖。
    /// </remarks>
    internal static async Task RunAsync(string output, bool asynchronous, IExcelExporter exporter)
    {
        output = Path.GetFullPath(output);
        if (File.Exists(output) || File.Exists(output + ".probe.json"))
            throw new ArgumentException("Use a new path for the advanced client fixture.");
        Directory.CreateDirectory(Path.GetDirectoryName(output));
        var rows = CreateRows();
        var headers = new[] { "Warehouse", "Category", "Month", "Status", "Inclusion", "Amount", "Item" };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows)
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "AdvancedWarehouse", SourceSheet = "Data", TargetSheet = "Summary",
                SourceRange = new ExcelRangeDefinition { EndRow = rows.Length, EndColumn = headers.Length - 1 },
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Warehouse", SortOrder = ExcelPivotSortOrder.Ascending },
                    new ExcelPivotAxisFieldDefinition { Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                        SubtotalMode = ExcelPivotSubtotalMode.None }
                },
                ColumnFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Month", SortOrder = ExcelPivotSortOrder.Ascending,
                        SubtotalMode = ExcelPivotSubtotalMode.None },
                    new ExcelPivotAxisFieldDefinition { Field = "Status", SortOrder = ExcelPivotSortOrder.Ascending,
                        SubtotalMode = ExcelPivotSubtotalMode.None }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", NumberFormat = "0.00" },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Mean", Aggregation = ExcelPivotAggregation.Average },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Low", Aggregation = ExcelPivotAggregation.Minimum },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "High", Aggregation = ExcelPivotAggregation.Maximum },
                    new ExcelPivotValueDefinition { Field = "Item", Caption = "Count", Aggregation = ExcelPivotAggregation.Count }
                },
                Filters = new[] { new ExcelPivotFilterDefinition { Field = "Inclusion", Values = new object[] { "yes", "Maybe" } } },
                MaxExpandedCells = 2000
            }));
        if (asynchronous)
            await exporter.ExportToFileAsync(request, output);
        else
            exporter.ExportToFile(request, output);

        using var workbook = new XLWorkbook(output);
        Require(workbook.Worksheets.Select(sheet => sheet.Name).SequenceEqual(new[] { "Data", "Summary" }),
            "Complete advanced worksheet order differs.");
        var expectedSource = headers.Concat(rows.SelectMany(Values)).ToArray();
        Require(workbook.Worksheet("Data").Range("A1:G8").Cells().Select(cell => cell.GetString())
            .SequenceEqual(expectedSource), "Complete advanced source differs.");
        var pivot = workbook.Worksheet("Summary").PivotTables.Single();
        Require(pivot.RowLabels.Select(field => field.SourceName).SequenceEqual(new[] { "Warehouse", "Category" }),
            "Advanced row roles differ.");
        Require(pivot.ColumnLabels.Where(field => field.SourceName != "{{Values}}")
            .Select(field => field.SourceName).SequenceEqual(new[] { "Month", "Status" }), "Advanced column roles differ.");
        Require(pivot.Values.Select(value => value.CustomName).SequenceEqual(new[] { "Total", "Mean", "Low", "High", "Count" }),
            "Advanced value titles differ.");
        Require(pivot.Values.Select(value => value.SummaryFormula).SequenceEqual(new[]
            { XLPivotSummary.Sum, XLPivotSummary.Average, XLPivotSummary.Minimum, XLPivotSummary.Maximum, XLPivotSummary.Count }),
            "Advanced aggregations differ.");
        Require(pivot.ReportFilters.Single().SelectedValues.Select(value => value.ToString())
            .SequenceEqual(new[] { "Yes", "Maybe" }), "Advanced selected filter values differ.");
        using var archive = ZipFile.OpenRead(output);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var cache = ReadPart(archive, "pivotCacheDefinition");
        var records = ReadPart(archive, "pivotCacheRecords");
        var fields = cache.Root.Element(ns + "cacheFields").Elements().ToArray();
        Require(fields.Select(field => (string)field.Attribute("name")).SequenceEqual(headers),
            "Complete advanced cache fields differ.");
        var source = cache.Root.Element(ns + "cacheSource").Element(ns + "worksheetSource");
        Require((string)source.Attribute("sheet") == "Data" && (string)source.Attribute("ref") == "A1:G8",
            "Advanced cache source differs.");
        var cachedRows = records.Root.Elements(ns + "r").ToArray();
        Require(cachedRows.Length == rows.Length, "Filtering deleted advanced cache records.");
        for (var row = 0; row < rows.Length; row++)
        {
            var cells = cachedRows[row].Elements().ToArray();
            Require(cells.Length == headers.Length, "Advanced cache record width differs.");
            var actual = cells.Select((cell, index) => cell.Name.LocalName == "x"
                ? (string)fields[index].Element(ns + "sharedItems").Elements()
                    .ElementAt(int.Parse((string)cell.Attribute("v"), CultureInfo.InvariantCulture)).Attribute("v")
                : (string)cell.Attribute("v"));
            Require(actual.SequenceEqual(Values(rows[row])), "Complete advanced cache record differs.");
        }
        var report = new
        {
            tool = "Bing.Offices.AdvancedPivotProbe", tfm = RuntimeInformation.FrameworkDescription,
            output, provider = exporter.GetType().Name, asynchronous, sourceRange = "A1:G8", expectedSource, sourceRows = rows,
            selectedInclusion = new[] { "Yes", "Maybe" },
            aggregations = new[] { "Total", "Mean", "Low", "High", "Count" },
            expectedGrandTotals = new[] { 150d, 30d, 10d, 50d, 5d },
            editAddress = "F2", editValue = 15d,
            expectedRefreshedGrandTotals = new[] { 155d, 31d, 15d, 50d, 5d },
            structuralVerification = "Passed", clientVerification = "Pending"
        };
        File.WriteAllText(output + ".probe.json", JsonSerializer.Serialize(report,
            new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n", StringComparison.Ordinal)
            + "\n", new UTF8Encoding(false));
        Console.WriteLine($"advanced-pivot-probe-ok tfm={RuntimeInformation.FrameworkDescription} async={asynchronous}");
    }

    /// <summary>
    /// 创建包含筛选排除项和重复聚合项的采购来源。
    /// </summary>
    /// <returns>固定的完整客户端验收来源。</returns>
    private static AdvancedPivotProbeRow[] CreateRows() => new[]
    {
        new AdvancedPivotProbeRow { Warehouse = "North", Category = "Food", Month = "Jan", Status = "Open", Inclusion = "Yes", Amount = 10, Item = "A" },
        new AdvancedPivotProbeRow { Warehouse = "North", Category = "Food", Month = "Feb", Status = "Closed", Inclusion = "No", Amount = 30, Item = "B" },
        new AdvancedPivotProbeRow { Warehouse = "North", Category = "Tools", Month = "Jan", Status = "Open", Inclusion = "Yes", Amount = 20, Item = "C" },
        new AdvancedPivotProbeRow { Warehouse = "South", Category = "Food", Month = "Jan", Status = "Pending", Inclusion = "Yes", Amount = 40, Item = "D" },
        new AdvancedPivotProbeRow { Warehouse = "North", Category = "Food", Month = "Jan", Status = "Open", Inclusion = "Yes", Amount = 30, Item = "E" },
        new AdvancedPivotProbeRow { Warehouse = "South", Category = "Tools", Month = "Feb", Status = "Pending", Inclusion = "Maybe", Amount = 50, Item = "F" },
        new AdvancedPivotProbeRow { Warehouse = "South", Category = "Tools", Month = "Feb", Status = "Closed", Inclusion = "No", Amount = 70, Item = "G" }
    };

    /// <summary>
    /// 将一条来源记录转为按列排列的验收值。
    /// </summary>
    /// <param name="row">采购来源记录。</param>
    /// <returns>与工作表列顺序一致的字面值。</returns>
    private static IEnumerable<string> Values(AdvancedPivotProbeRow row) => new[]
    {
        row.Warehouse, row.Category, row.Month, row.Status, row.Inclusion,
        row.Amount.ToString(CultureInfo.InvariantCulture), row.Item
    };

    /// <summary>
    /// 读取唯一透视部件。
    /// </summary>
    /// <param name="archive">真实 XLSX 包。</param>
    /// <param name="prefix">XML 部件文件名前缀。</param>
    /// <returns>完整的透视部件文档。</returns>
    private static XDocument ReadPart(ZipArchive archive, string prefix)
    {
        var entry = archive.Entries.Single(item => item.FullName.Split('/').Last()
            .StartsWith(prefix, StringComparison.Ordinal) && item.FullName.EndsWith(".xml", StringComparison.Ordinal));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 检查客户端夹具的必要条件。
    /// </summary>
    /// <param name="condition">是否符合预期。</param>
    /// <param name="message">失败原因。</param>
    private static void Require(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }
}
