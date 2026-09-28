using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Profiles;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证实体列表分页小计的 Provider 公共合同。
/// </summary>
public sealed class EntityPageSubtotalContractTest
{
    /// <summary>
    /// NPOI 和 ClosedXML 应写入并跳过分页小计，MiniExcel 不适用。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("MiniExcel")]
    [InlineData("ClosedXML")]
    public void EntityPageSubtotals_ShouldMatchIndependentProfile(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        var expectation = ProviderContractProfiles.Get(provider).For(ContractScenario.EntityPageSubtotals);

        if (expectation == ContractExpectation.NotApplicable)
        {
            Assert.Equal("MiniExcel", provider);
            Assert.False(driver.DeclaredCapabilities.Supports(ExcelProviderCapabilities.Entity));
            return;
        }

        Assert.Equal(ContractExpectation.Supported, expectation);
        var layout = CreateLayout();
        var source = new EntityPageSubtotalContractRoot
        {
            Lines = new List<EntityPageSubtotalContractLine>
            {
                new() { Category = "A", Amount = 2m },
                new() { Category = "B", Amount = 3m },
                new() { Category = "C", Amount = 4m }
            }
        };

        using var output = new MemoryStream();
        var exporter = Assert.IsAssignableFrom<IExcelEntityExporter>(driver.CreateExporter());
        exporter.ExportEntity(source, layout, output);
        Assert.NotEmpty(output.ToArray());

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var importer = Assert.IsAssignableFrom<IExcelEntityImporter>(driver.CreateImporter());
        var result = importer.ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B", "C" }, result.Entity.Lines.Select(line => line.Category).ToArray());
        Assert.Equal(new[] { 2m, 3m, 4m }, result.Entity.Lines.Select(line => line.Amount).ToArray());
    }

    /// <summary>
    /// 创建包含分页小计和最终尾部的实体布局。
    /// </summary>
    /// <returns>包含分页小计和最终合计的实体布局。</returns>
    private static ExcelEntityLayout<EntityPageSubtotalContractRoot> CreateLayout() =>
        ExcelEntity.Layout<EntityPageSubtotalContractRoot>(builder => builder
            .ListRegion("Data", "A1", root => root.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))
                .Footer("TOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))));

    /// <summary>
    /// 分页小计合同的根聚合模型。
    /// </summary>
    private sealed class EntityPageSubtotalContractRoot
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<EntityPageSubtotalContractLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 分页小计合同的明细行模型。
    /// </summary>
    private sealed class EntityPageSubtotalContractLine
    {
        /// <summary>
        /// 获取或设置明细名称。
        /// </summary>
        public string Category { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
