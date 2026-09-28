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
/// 验证实体列表连续分组小计的 Provider 公共合同。
/// </summary>
public sealed class EntityGroupSubtotalContractTest
{
    /// <summary>
    /// NPOI 和 ClosedXML 应跳过小计行，MiniExcel 不适用。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("MiniExcel")]
    [InlineData("ClosedXML")]
    public void EntityGroupSubtotals_ShouldMatchIndependentProfile(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        var expectation = ProviderContractProfiles.Get(provider).For(ContractScenario.EntityGroupSubtotals);

        if (expectation == ContractExpectation.NotApplicable)
        {
            Assert.Equal("MiniExcel", provider);
            Assert.False(driver.DeclaredCapabilities.Supports(ExcelProviderCapabilities.Entity));
            return;
        }

        Assert.Equal(ContractExpectation.Supported, expectation);
        var layout = CreateLayout();
        var source = new EntityGroupSubtotalContractRoot
        {
            Lines = new List<EntityGroupSubtotalContractLine>
            {
                new() { Category = "A", Amount = 2m },
                new() { Category = "A", Amount = 3m },
                new() { Category = "B", Amount = 4m }
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
        Assert.Equal(new[] { "A", "A", "B" }, result.Entity.Lines.Select(line => line.Category).ToArray());
        Assert.Equal(new[] { 2m, 3m, 4m }, result.Entity.Lines.Select(line => line.Amount).ToArray());
    }

    /// <summary>
    /// 创建包含连续分组小计和最终尾部的实体布局。
    /// </summary>
    /// <returns>包含小计和最终合计的实体布局。</returns>
    private static ExcelEntityLayout<EntityGroupSubtotalContractRoot> CreateLayout() =>
        ExcelEntity.Layout<EntityGroupSubtotalContractRoot>(builder => builder
            .ListRegion("Data", "A1", root => root.Lines, region => region
                .GroupSubtotal(line => line.Category, "SUBTOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))
                .Footer("TOTAL", footer => footer
                    .Cell("B1", lines => lines.Sum(line => line.Amount), numberFormat: "0.00"))));

    /// <summary>
    /// 连续分组小计合同的根聚合模型。
    /// </summary>
    private sealed class EntityGroupSubtotalContractRoot
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<EntityGroupSubtotalContractLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 连续分组小计合同的明细行模型。
    /// </summary>
    private sealed class EntityGroupSubtotalContractLine
    {
        /// <summary>
        /// 获取或设置分组名称。
        /// </summary>
        public string Category { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
