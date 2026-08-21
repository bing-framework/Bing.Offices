using System.Collections.Generic;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Docs.Tests;

/// <summary>
/// 验证文档消费者只能通过 provider-neutral 接口使用 NPOI 注册入口。
/// </summary>
public sealed class DocsConsumerTest
{
    /// <summary>
    /// 测试 - 外部消费者应能调用 AddNpoi 并解析导入导出接口。
    /// </summary>
    [Fact]
    public void AddNpoi_ExternalConsumer_ShouldResolveProviderNeutralServices()
    {
        // Arrange
        var services = new ServiceCollection();

        // Act
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();

        // Assert
        Assert.NotNull(provider.GetRequiredService<IExcelImporter>());
        Assert.NotNull(provider.GetRequiredService<IExcelExporter>());
    }

    /// <summary>
    /// 测试 - 文档请求应只构造 provider-neutral Workbook API。
    /// </summary>
    [Fact]
    public void WorkbookRequest_ExternalConsumer_ShouldBuildWithoutNpoiTypes()
    {
        // Act
        var request = ExcelImport.Workbook<DocsWorkbook>(builder =>
            builder.Sheet("Data", workbook => workbook.Rows));
        var export = ExcelExport.Workbook(builder =>
            builder.AddSheet("Data", new[] { new DocsRow { Name = "consumer" } }));

        // Assert
        Assert.Equal(1, request.SheetCount);
        Assert.Equal(1, export.SheetCount);
    }

    private sealed class DocsWorkbook
    {
        public List<DocsRow> Rows { get; } = new();
    }

    private sealed class DocsRow
    {
        public string Name { get; set; }
    }
}
