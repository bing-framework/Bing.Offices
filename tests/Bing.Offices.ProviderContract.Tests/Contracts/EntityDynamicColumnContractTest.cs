using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Profiles;
using Bing.Offices.Testing.Models;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证实体动态列的 Provider profile、布局和真实 XLSX 往返。
/// </summary>
public sealed class EntityDynamicColumnContractTest
{
    /// <summary>
    /// ClosedXML 执行动态列合同，NPOI 显式 Unsupported，MiniExcel 不适用。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("MiniExcel")]
    [InlineData("ClosedXML")]
    public void EntityDynamicColumns_ShouldMatchIndependentProfile(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        var expectation = ProviderContractProfiles.Get(provider).For(ContractScenario.EntityDynamicColumns);
        var layout = CreateLayout();
        var entity = new EntityDynamicContractRoot
        {
            Lines = new List<EntityDynamicContractLine>
            {
                new()
                {
                    Name = "A",
                    Quantity = 2,
                    CustomFields = new Dictionary<string, object>(System.StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };

        if (expectation == ContractExpectation.NotApplicable)
        {
            Assert.Equal("MiniExcel", provider);
            Assert.False(driver.DeclaredCapabilities.Supports(Bing.Offices.Providers.ExcelProviderCapabilities.Entity));
            return;
        }

        using var output = new MemoryStream();
        var exporter = Assert.IsAssignableFrom<IExcelEntityExporter>(driver.CreateExporter());
        if (expectation == ContractExpectation.UnsupportedExpected)
        {
            var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                exporter.ExportEntity(entity, layout, output));
            Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
            Assert.Empty(output.ToArray());
            return;
        }

        exporter.ExportEntity(entity, layout, output);
        Assert.NotEmpty(output.ToArray());
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = Assert.IsAssignableFrom<IExcelEntityImporter>(driver.CreateImporter())
            .ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("A", line.Name);
        Assert.Equal(2, line.Quantity);
        Assert.Equal("North", line.CustomFields["zone"]);
    }

    /// <summary>
    /// 创建包含固定列和动态区域列的实体布局。
    /// </summary>
    /// <returns>包含动态 zone 列的实体布局。</returns>
    private static ExcelEntityLayout<EntityDynamicContractRoot> CreateLayout() =>
        ExcelEntity.Layout<EntityDynamicContractRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines, region => region.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(EntityDynamicContractLine.Name), Title = "Item" },
                        new() { PropertyName = nameof(EntityDynamicContractLine.Quantity), Title = "Qty" }
                    },
                    DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                    {
                        new()
                        {
                            Key = "zone", Title = "Zone", DataTypeName = "string", Order = 0,
                            PlacementKey = $"before:{nameof(EntityDynamicContractLine.Quantity)}"
                        }
                    }
                })));
}
