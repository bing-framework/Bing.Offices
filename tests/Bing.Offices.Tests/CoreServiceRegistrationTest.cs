using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Extensions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>Core-only 注册的真实 CSV 行为、幂等性和预注册实现优先级。</summary>
public sealed class CoreServiceRegistrationTest
{
    /// <summary>不注册 Excel Provider 时 CSV 导入、导出和默认校验必须独立可用。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AddBingOfficesCore_ShouldSupportCsvWithoutExcelProvider(bool asynchronous)
    {
        var services = new ServiceCollection();
        Assert.Same(services, services.AddBingOfficesCore());
        using var provider = services.BuildServiceProvider();
        Assert.Null(provider.GetService<IExcelImporter>());
        Assert.Null(provider.GetService<IExcelExporter>());
        Assert.NotNull(provider.GetRequiredService<IExcelMappingConfigurationLoader>());
        Assert.NotNull(provider.GetRequiredService<IFileExportCommitter>());
        var exporter = provider.GetRequiredService<ICsvExporter>();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var data = new[] { new Row { Code = "ORD-123" } };
        var bytes = asynchronous ? await exporter.ExportToBytesAsync(data) : exporter.ExportToBytes(data);
        Assert.Equal("Code\r\nORD-123\r\n", Encoding.UTF8.GetString(bytes));
        var result = asynchronous ? await importer.ImportFromBytesAsync<Row>(bytes) : importer.ImportFromBytes<Row>(bytes);
        Assert.Equal("ORD-123", Assert.Single(result.Items).Code);
        Assert.Empty(result.Errors);
        Assert.True(result.IsSuccess);
        Assert.True(result.IsComplete);
        var invalid = Encoding.UTF8.GetBytes("Code\r\ninvalid\r\nORD-123\r\n");
        var failure = asynchronous ? await importer.ImportFromBytesAsync<Row>(invalid) : importer.ImportFromBytes<Row>(invalid);
        Assert.Equal("ORD-123", Assert.Single(failure.Items).Code);
        var error = Assert.Single(failure.Errors);
        Assert.Equal(CsvImportErrorCode.Validation, error.Code);
        Assert.Equal("格式不正确", error.Message);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(Row.Code), error.PropertyName);
        Assert.Null(error.FirstRowNumber);
        Assert.False(failure.IsSuccess);
        Assert.True(failure.IsComplete);
    }

    /// <summary>Core 与 NPOI 任意重复注册顺序都不应重复默认规则，也不替换用户已注册服务。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddBingOfficesCore_ShouldBeIdempotentAndPreserveExistingServices(bool npoiFirst)
    {
        using var existing = new ServiceCollection().AddBingOfficesCore().BuildServiceProvider();
        var importer = existing.GetRequiredService<ICsvImporter>();
        var exporter = existing.GetRequiredService<ICsvExporter>();
        var factory = existing.GetRequiredService<IExcelMappingPlanFactory>();
        var committer = existing.GetRequiredService<IFileExportCommitter>();
        var rule = existing.GetServices<IExcelValidationRule>().First();
        var services = new ServiceCollection();
        services.AddSingleton(importer).AddSingleton(exporter).AddSingleton(factory).AddSingleton(committer);
        services.AddSingleton<IExcelValidationRule>(rule);
        if (npoiFirst)
            services.AddBingOfficesNpoi();
        services.AddBingOfficesCore().AddBingOfficesCore();
        services.AddBingOfficesNpoi().AddBingOfficesCore();
        using var provider = services.BuildServiceProvider();
        Assert.Same(importer, provider.GetRequiredService<ICsvImporter>());
        Assert.Same(exporter, provider.GetRequiredService<ICsvExporter>());
        Assert.Same(factory, provider.GetRequiredService<IExcelMappingPlanFactory>());
        Assert.Same(committer, provider.GetRequiredService<IFileExportCommitter>());
        Assert.NotNull(provider.GetRequiredService<IExcelImporter>());
        Assert.NotNull(provider.GetRequiredService<IExcelExporter>());
        var rules = provider.GetServices<IExcelValidationRule>().ToArray();
        Assert.Contains(rule, rules);
        Assert.Equal(existing.GetServices<IExcelValidationRule>().Select(item => item.GetType()).OrderBy(type => type.FullName),
            rules.Select(item => item.GetType()).OrderBy(type => type.FullName));
        Assert.Equal(rules.Length, rules.Select(item => item.GetType()).Distinct().Count());
    }

    /// <summary>空服务集合应在入口处准确报告参数名。</summary>
    [Fact]
    public void AddBingOfficesCore_NullServices_ShouldThrow()
        => Assert.Equal("services", Assert.Throws<ArgumentNullException>(() =>
            BingOfficesServiceCollectionExtensions.AddBingOfficesCore(null)).ParamName);

    private sealed class Row
    {
        [ExcelRegex(@"^ORD-\d+$")]
        public string Code { get; set; }
    }
}
