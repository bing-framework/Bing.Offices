using System.IO;
using Bing.Offices.Formula;
using Bing.Offices.Npoi.Formula;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 公式文本和缓存值合同。
/// </summary>
public sealed class NpoiFormulaContractTest
{
    /// <summary>
    /// 验证公式与缓存值读取模式保留公式文本。
    /// </summary>
    [Fact]
    public void Process_ShouldReadFormulaTextAndCachedValue()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue(1d);
        sheet.GetRow(0).CreateCell(1).SetCellFormula("A1*2");
        using var source = new MemoryStream();
        workbook.Write(source, true);
        source.Position = 0;

        var result = new NpoiExcelFormulaProcessor().Process(source, Bing.Offices.ExcelFormat.Xlsx,
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
            new NpoiExcelFormulaProcessor().Process(source, Bing.Offices.ExcelFormat.Xlsx,
                new ExcelFormulaRequest { CalculationMode = ExcelFormulaCalculationMode.MustRecalculate }));

        Assert.Equal(Bing.Offices.Exceptions.BingOfficesStage.Preflight, exception.Stage);
    }
}
