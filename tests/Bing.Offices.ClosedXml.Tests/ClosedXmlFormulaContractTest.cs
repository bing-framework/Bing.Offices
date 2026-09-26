using System.IO;
using Bing.Offices.ClosedXml.Formula;
using Bing.Offices.Formula;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML 公式文本和缓存值合同。
/// </summary>
public sealed class ClosedXmlFormulaContractTest
{
    /// <summary>
    /// 验证公式与缓存值读取模式保留公式文本。
    /// </summary>
    [Fact]
    public void Process_ShouldReadFormulaTextAndCachedValue()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Data");
        sheet.Cell("A1").Value = 1;
        sheet.Cell("B1").FormulaA1 = "A1*2";
        using var source = new MemoryStream();
        workbook.SaveAs(source);
        source.Position = 0;

        var result = new ClosedXmlExcelFormulaProcessor().Process(source, Bing.Offices.ExcelFormat.Xlsx,
            new ExcelFormulaRequest { ReadMode = ExcelFormulaReadMode.FormulaAndCachedValue });

        var cell = Assert.Single(result.Cells);
        Assert.Equal("A1*2", cell.Formula);
        Assert.True(result.CalculationCompleted);
    }

    /// <summary>
    /// 重新计算请求必须显式拒绝。
    /// </summary>
    [Fact]
    public void Process_Recalculate_ShouldBeUnsupported()
    {
        using var source = new MemoryStream();
        var exception = Assert.Throws<Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelFormulaProcessor().Process(source, Bing.Offices.ExcelFormat.Xlsx,
                new ExcelFormulaRequest { CalculationMode = ExcelFormulaCalculationMode.MustRecalculate }));

        Assert.Equal(Bing.Offices.Exceptions.BingOfficesStage.Preflight, exception.Stage);
    }
}
