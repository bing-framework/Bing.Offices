using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.Exports;
using Xunit;
using Xunit.Abstractions;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证透视请求在大型列表阈值仍可读取完整来源。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotStreamingBoundaryTest
{
    /// <summary>
    /// 输出本次独立边界探针的时间与进程资源事实。
    /// </summary>
    private readonly ITestOutputHelper _output;

    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotStreamingBoundaryTest" /> 类型的实例。
    /// </summary>
    /// <param name="output">测试运行器提供的结果输出。</param>
    public NpoiPivotStreamingBoundaryTest(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// 验证透视保留全部业务行且只缓存明确声明的来源区域。
    /// </summary>
    /// <remarks>
    /// 本测试只在计划要求的 100K 定向探针中执行，常规门禁排除 Large。
    /// 它验证 SXSSF 阈值上的执行选择，不声明恒定内存或性能目标。
    /// </remarks>
    [Fact]
    [Trait("Category", "Large")]
    public void Export_AtStreamingThreshold_ShouldKeepCompleteRowsAndBoundedPivotSource()
    {
        const int count = 100000;
        var timer = Stopwatch.StartNew();
        var rows = Enumerable.Range(0, count).Select(index => new NpoiPivotBoundaryRow
        { Region = "Warehouse-" + index % 3, Amount = index + 0.25 }).ToArray();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows)
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "ThresholdPivot", SourceSheet = "Data", TargetSheet = "Summary",
                SourceRange = new ExcelRangeDefinition { EndRow = 3, EndColumn = 1 },
                RowField = "Region", ValueField = "Amount"
            }));
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        Assert.True(output.CanWrite);
        output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var strings = ReadXml(archive, "xl/sharedStrings.xml").Root.Elements(ns + "si")
            .Select(item => string.Concat(item.Descendants(ns + "t").Select(text => text.Value))).ToArray();
        using (var input = archive.GetEntry("xl/worksheets/sheet1.xml").Open())
        using (var reader = XmlReader.Create(input, new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
        {
            var rowIndex = 0;
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row")
                    continue;
                using var subtree = reader.ReadSubtree();
                var row = XElement.Load(subtree);
                Assert.Equal(rowIndex + 1, (int)row.Attribute("r"));
                var cells = row.Elements(ns + "c").ToArray();
                Assert.Equal(2, cells.Length);
                Assert.Equal("A" + (rowIndex + 1), (string)cells[0].Attribute("r"));
                Assert.Equal("B" + (rowIndex + 1), (string)cells[1].Attribute("r"));
                Assert.Equal("s", (string)cells[0].Attribute("t"));
                var first = strings[(int)cells[0].Element(ns + "v")];
                if (rowIndex == 0)
                {
                    Assert.Equal("Region", first);
                    Assert.Equal("Amount", strings[(int)cells[1].Element(ns + "v")]);
                }
                else
                {
                    Assert.Equal(rows[rowIndex - 1].Region, first);
                    Assert.Equal(rows[rowIndex - 1].Amount,
                        double.Parse(cells[1].Element(ns + "v").Value, CultureInfo.InvariantCulture));
                }
                rowIndex++;
            }
            Assert.Equal(count + 1, rowIndex);
        }
        var cachePath = archive.Entries.Single(entry => Path.GetFileName(entry.FullName)
            .StartsWith("pivotCacheDefinition", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var recordsPath = archive.Entries.Single(entry => Path.GetFileName(entry.FullName)
            .StartsWith("pivotCacheRecords", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var cache = ReadXml(archive, cachePath).Root;
        Assert.Equal(3, (int)cache.Attribute("recordCount"));
        Assert.Equal("A1:B4", (string)cache.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref"));
        var fields = cache.Element(ns + "cacheFields").Elements().ToArray();
        Assert.Equal(new[] { "Region", "Amount" }, fields.Select(field => (string)field.Attribute("name")));
        var records = ReadXml(archive, recordsPath).Root.Elements(ns + "r").ToArray();
        Assert.Equal(3, records.Length);
        for (var row = 0; row < records.Length; row++)
        {
            var indexes = records[row].Elements(ns + "x").Select(item => (int)item.Attribute("v")).ToArray();
            Assert.Equal(2, indexes.Length);
            Assert.Equal(rows[row].Region,
                (string)fields[0].Element(ns + "sharedItems").Elements().ElementAt(indexes[0]).Attribute("v"));
            Assert.Equal(rows[row].Amount,
                (double)fields[1].Element(ns + "sharedItems").Elements().ElementAt(indexes[1]).Attribute("v"));
        }
        _output.WriteLine("rows={0};cells={1};cachedRows={2};serializedBytes={3};elapsedMs={4};peakWorkingSet={5}",
            count, (count + 1) * 2, records.Length, output.Length, timer.ElapsedMilliseconds,
            Process.GetCurrentProcess().PeakWorkingSet64);
    }

    /// <summary>
    /// 读取指定工作簿 XML 部件。
    /// </summary>
    /// <param name="archive">真实导出的工作簿包。</param>
    /// <param name="path">部件在包内的地址。</param>
    /// <returns>完整 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string path)
    {
        using var input = archive.GetEntry(path).Open();
        return XDocument.Load(input);
    }
}
