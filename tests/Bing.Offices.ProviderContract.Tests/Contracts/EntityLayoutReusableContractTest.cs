using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Entities;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 NPOI 与 ClosedXML 可复用 Entity Layout Provider 合同。
/// </summary>
public sealed class EntityLayoutReusableContractTest
{
    /// <summary>
    /// 运行核心 Entity Layout 往返合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void Core_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyCore(exporter, importer);
    }

    /// <summary>
    /// 运行分页小计 Entity Layout 合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void PageSubtotal_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyPageSubtotal(exporter, importer);
    }

    /// <summary>运行重复页眉 Entity Layout 合同。</summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void PageHeaders_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyPageHeaders(exporter, importer);
    }

    /// <summary>验证重复页眉工作簿的输入字节限制。</summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void PageHeaderInputLimit_ShouldRejectBeforeRead(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);
        var layout = ExcelEntity.Layout<DetailSumOrder>(builder => builder
            .ListRegion("Data", "B3", order => order.Lines, region => region
                .PageBreak(1).PageHeader("A1:B3", "A1", "HEAD")));
        var source = new DetailSumOrder
        {
            Lines = new List<DetailSumLine>
            {
                new() { Amount = 2m },
                new() { Amount = 3m }
            }
        };
        using var output = new MemoryStream();
        exporter.ExportEntity(source, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = input.Length - 1
        });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
        {
            if (provider == "NPOI")
                ((NpoiExcelImporter)importer).ImportEntity(input, layout, options);
            else
                ((ClosedXmlExcelImporter)importer).ImportEntity(input, layout, options);
        });
        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(provider == "NPOI" ? BingOfficesStage.Open : BingOfficesStage.Preflight,
            exception.Stage);
        Assert.True(input.CanRead);
    }

    /// <summary>
    /// 运行连续分组小计 Entity Layout 合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void GroupSubtotal_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyGroupSubtotal(exporter, importer);
    }

    /// <summary>
    /// 运行尾部原生公式与连续明细求和合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void FooterFormulas_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyFooterFormulas(exporter, importer);
    }

    /// <summary>
    /// 运行 Entity Layout 异步导入导出合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task Async_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        await EntityLayoutProviderContractSuite.VerifyAsync(exporter, importer);
    }

    /// <summary>
    /// 运行属性式固定单元格与命名锚点模板往返合同。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void AttributeAndNamedAnchors_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyAttributeAndNamedAnchors(exporter, importer);
    }

    /// <summary>
    /// 命名列表与固定单元格按模板实际地址重叠时，导入导出均应返回配置错误。
    /// </summary>
    /// <param name="provider">待验证的 XLSX Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void NamedListFixedCellCollision_ShouldSatisfyReusableContract(string provider)
    {
        var (exporter, importer) = CreateEntityServices(provider);

        EntityLayoutProviderContractSuite.VerifyNamedListFixedCellCollision(exporter, importer);
    }

    /// <summary>
    /// 确认合同入口对缺失的 Provider 服务给出明确失败。
    /// </summary>
    [Fact]
    public void Core_ShouldRejectMissingServices()
    {
        Assert.Throws<ArgumentNullException>(() =>
            EntityLayoutProviderContractSuite.VerifyCore(null, null));
    }

    /// <summary>
    /// 验证跨小计总计公式只引用有序明细行段。
    /// </summary>
    [Fact]
    public void DetailSumFormula_ShouldValidateAndPartitionDetailRanges()
    {
        var layout = ExcelEntity.Layout<DetailSumOrder>(builder => builder
            .ListRegion("Data", "A1", order => order.Lines, region => region
                .Footer("TOTAL", footer => footer.FormulaSumDetailRowsAbove("C1", " b "))));
        var cell = Assert.Single(Assert.Single(layout.ListRegions).Footer.Cells);
        var formula = Assert.IsType<ExcelEntityFooterDetailSumValue>(
            cell.Evaluate(Array.Empty<object>()));

        Assert.Equal("B", formula.SourceColumn);
        Assert.Equal("0", formula.ToFormulaA1(Array.Empty<(int, int)>()));
        Assert.Equal("SUM(B1:B2,B6:B8)", formula.ToFormulaA1(new[] { (1, 2), (6, 8) }));
        Assert.Throws<ArgumentException>(() => formula.ToFormulaA1(new[] { (1, 2), (2, 3) }));
        Assert.Throws<ArgumentException>(() => formula.ToFormulaA1(new[] { (3, 2) }));
        var manyRanges = Enumerable.Range(1, 256).Select(index => (index * 2, index * 2))
            .ToArray();
        var partitioned = formula.ToFormulaA1(manyRanges);
        Assert.Equal(2, partitioned.Split(new[] { "SUM(" }, StringSplitOptions.None).Length - 1);
        Assert.Contains(")+SUM(", partitioned);
        Assert.Throws<ArgumentException>(() => formula.ToFormulaA1(Enumerable.Range(1, 1000)
            .Select(index => (index * 2, index * 2)).ToArray()));
    }

    /// <summary>
    /// 创建指定 Provider 的公开实体导入及导出服务。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <returns>实体导出器和导入器。</returns>
    private static (IExcelEntityExporter Exporter, IExcelEntityImporter Importer)
        CreateEntityServices(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        return (
            Assert.IsAssignableFrom<IExcelEntityExporter>(driver.CreateExporter()),
            Assert.IsAssignableFrom<IExcelEntityImporter>(driver.CreateImporter()));
    }

    /// <summary>
    /// 跨小计公式测试的根实体。
    /// </summary>
    private sealed class DetailSumOrder
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<DetailSumLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 跨小计公式测试的明细实体。
    /// </summary>
    private sealed class DetailSumLine
    {
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
