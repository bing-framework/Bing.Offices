using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Profiles;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 通过公共导入 API 验证冻结公式输入的公式文本保留行为。
/// </summary>
public sealed class FormulaGoldenContractTest
{
    /// <summary>
    /// 三个 Provider 都应对冻结公式输入给出独立档案声明的结果。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <param name="expectedValue">按 Provider 契约预期导入的文本值；允许为 null。</param>
    [Theory]
    [InlineData("NPOI", "formula/cached-number.xlsx", "2")]
    [InlineData("NPOI", "formula/cached-bool.xlsx", "True")]
    [InlineData("NPOI", "formula/cached-string.xlsx", "AB")]
    [InlineData("NPOI", "formula/no-cache.xlsx", "0")]
    [InlineData("MiniExcel", "formula/cached-number.xlsx", "2")]
    [InlineData("MiniExcel", "formula/cached-bool.xlsx", "True")]
    [InlineData("MiniExcel", "formula/cached-string.xlsx", "AB")]
    [InlineData("MiniExcel", "formula/no-cache.xlsx", null)]
    [InlineData("ClosedXML", "formula/cached-number.xlsx", "=1+1")]
    [InlineData("ClosedXML", "formula/cached-bool.xlsx", "=1=1")]
    [InlineData("ClosedXML", "formula/cached-string.xlsx", "=CONCAT(\"A\",\"B\")")]
    [InlineData("ClosedXML", "formula/no-cache.xlsx", "=1+1")]
    public async Task FormulaGolden_ShouldMatchIndependentProviderProfile(
        string provider, string relativePath, string expectedValue)
    {
        using var source = Golden(relativePath);
        var request = ExcelImport.Workbook<FormulaGoldenWorkbook>(workbook =>
            workbook.Sheet<FormulaGoldenRow>("Formula", root => root.Rows));

        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var expectation = ProviderContractProfiles.Get(provider).For(ContractScenario.Formula);
        Assert.Equal(provider == "ClosedXML" ? ContractExpectation.Supported : ContractExpectation.Partial,
            expectation);
        Assert.Equal(expectedValue, Assert.Single(result.Workbook.Rows).Value);
    }

    /// <summary>
    /// 读取黄金样例文件内容。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <returns>包含黄金样例内容的只读内存流。</returns>
    private static MemoryStream Golden(string relativePath)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Resources", "Golden",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(path), $"Golden input was not copied: {path}");
        return new MemoryStream(File.ReadAllBytes(path), writable: false);
    }

    /// <summary>
    /// 承载公式黄金样例导入结果的测试工作簿。
    /// </summary>
    private sealed class FormulaGoldenWorkbook
    {
        /// <summary>
        /// 获取公式样例导入的数据行集合。
        /// </summary>
        public System.Collections.Generic.List<FormulaGoldenRow> Rows { get; } = new();
    }

    /// <summary>
    /// 用于比较公式导入结果的测试数据行。
    /// </summary>
    private sealed class FormulaGoldenRow
    {
        /// <summary>
        /// 获取或设置导入的公式文本或缓存值文本。
        /// </summary>
        public string Value { get; set; } = string.Empty;
    }
}
