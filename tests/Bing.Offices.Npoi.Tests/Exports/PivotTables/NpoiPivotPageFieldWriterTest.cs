using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Exports.PivotTables;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证透视单选索引零的快照与实际 XML 补齐。
/// </summary>
[Collection("NpoiNativePivot")]
public sealed class NpoiPivotPageFieldWriterTest
{
    /// <summary>
    /// 验证明确单选零与未指定全选的快照保持隔离。
    /// </summary>
    [Fact]
    public void Capture_ShouldDistinguishZeroSelectionFromAllAndRemainImmutable()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var header = sheet.CreateRow(0);
        foreach (var pair in new[] { "Region", "Amount", "State" }.Select((name, index) => (name, index)))
            header.CreateCell(pair.index).SetCellValue(pair.name);
        for (var index = 0; index < 2; index++)
        {
            var row = sheet.CreateRow(index + 1);
            row.CreateCell(0).SetCellValue(index == 0 ? "North" : "South");
            row.CreateCell(1).SetCellValue(10d + index);
            row.CreateCell(2).SetCellValue(index == 0 ? "Open" : "Closed");
        }
        var definition = new ExcelPivotTableDefinition
        {
            Name = "SelectedZero", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { EndRow = 2, EndColumn = 2 },
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
            Filters = new[] { new ExcelPivotFilterDefinition { Field = "State", Values = new object[] { "Open" } } }
        };
        var source = NpoiPivotSourceValidator.Validate(workbook, definition, default);
        NpoiPivotTableWriter.Create(workbook, source, default);
        var pivot = ((XSSFSheet)workbook.GetSheet("Summary")).GetPivotTables().Single();
        var page = pivot.GetCTPivotTableDefinition().pageFields.pageField.Single();
        Assert.True(page.itemSpecified);
        Assert.Equal(0u, page.item);
        var snapshot = NpoiPivotPageFieldWriter.Capture(workbook, default);
        var selection = Assert.Single(snapshot);
        Assert.Equal(pivot.GetPackagePart().PartName.Name.TrimStart('/'), selection.Key);
        Assert.Equal(0u, selection.Value[2]);
        Assert.Throws<NotSupportedException>(() => ((IDictionary<int, uint>)selection.Value).Add(7, 0));
        page.item = 1;
        Assert.Equal(0u, selection.Value[2]);
        Assert.Empty(NpoiPivotPageFieldWriter.Capture(workbook, default));
        page.item = 0;
        page.itemSpecified = false;
        Assert.Empty(NpoiPivotPageFieldWriter.Capture(workbook, default));
    }

    /// <summary>
    /// 验证补齐零索引保留非零单选、全选及全部其他定义。
    /// </summary>
    [Fact]
    public void Write_ShouldPreserveCompleteDefinitionAndLeaveStreamsOpen()
    {
        const string xml = "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"Purchase\" cacheId=\"1\"><location ref=\"A5:D9\" firstHeaderRow=\"0\" firstDataRow=\"1\" firstDataCol=\"1\"/><pageFields count=\"3\"><pageField fld=\"2\" hier=\"-1\"/><pageField fld=\"3\" item=\"2\" hier=\"-1\"/><pageField fld=\"4\" hier=\"-1\"/></pageFields><pivotTableStyleInfo name=\"PivotStyleLight16\" showRowHeaders=\"1\"/></pivotTableDefinition>";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using var output = new MemoryStream();
        NpoiPivotPageFieldWriter.Write(source, output, new Dictionary<int, uint> { [2] = 0 }, default);
        Assert.True(source.CanRead);
        Assert.True(output.CanWrite);
        output.Position = 0;
        var actual = XDocument.Load(output);
        var expected = XDocument.Parse(xml);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        expected.Root.Element(ns + "pageFields").Elements().First().SetAttributeValue("item", 0);
        Assert.True(XNode.DeepEquals(expected, actual), actual.ToString());
    }

    /// <summary>
    /// 验证预取消和缺失字段均在写入前保全目标内容。
    /// </summary>
    [Fact]
    public void Write_CancellationAndMissingField_ShouldPreserveDestination()
    {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(
            "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><pageFields count=\"1\"><pageField fld=\"2\"/></pageFields></pivotTableDefinition>"));
        using var output = new MemoryStream();
        var original = new byte[] { 17, 42, 91 };
        output.Write(original);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var canceled = Assert.ThrowsAny<OperationCanceledException>(() =>
            NpoiPivotPageFieldWriter.Write(source, output, new Dictionary<int, uint> { [2] = 0 }, cancellation.Token));
        Assert.Equal(cancellation.Token, canceled.CancellationToken);
        Assert.Equal(original, output.ToArray());
        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            NpoiPivotPageFieldWriter.Write(source, output, new Dictionary<int, uint> { [7] = 0 }, default));
        Assert.Equal(BingOfficesStage.Write, error.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
    }
}
