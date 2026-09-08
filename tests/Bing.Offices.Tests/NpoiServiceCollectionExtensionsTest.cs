using System;
using System.Linq;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 服务注册扩展的职责级测试。
/// </summary>
public class NpoiServiceCollectionExtensionsTest
{
    /// <summary>
    /// 注册入口应返回原集合，并对空集合提供稳定的参数异常。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_ShouldPreserveChainAndValidateServices()
    {
        var services = new ServiceCollection();

        var actual = ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(services);
        var exception = Assert.Throws<ArgumentNullException>(() =>
            ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(null));

        Assert.Same(services, actual);
        Assert.Equal("services", exception.ParamName);
    }

    /// <summary>
    /// 重复调用不应重复注册单值服务或默认校验规则。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_WhenRepeated_ShouldRemainIdempotent()
    {
        var services = new ServiceCollection();

        ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(services);
        var first = services
            .GroupBy(descriptor => (descriptor.ServiceType, descriptor.ImplementationType))
            .ToDictionary(group => group.Key, group => group.Count());
        ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(services);
        var second = services
            .GroupBy(descriptor => (descriptor.ServiceType, descriptor.ImplementationType))
            .ToDictionary(group => group.Key, group => group.Count());

        Assert.Equal(first, second);
    }

    /// <summary>
    /// 调用方预注册的导入导出器应优先，默认请求服务保持 transient。
    /// </summary>
    [Fact]
    public void AddBingOfficesNpoi_ShouldPreserveReplacementsAndLifetimes()
    {
        var importer = new ReplacementImporter();
        var exporter = new ReplacementExporter();
        var services = new ServiceCollection();
        services.AddSingleton<IExcelImporter>(importer);
        services.AddSingleton<IExcelExporter>(exporter);

        ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi(services);
        using var provider = services.BuildServiceProvider();

        Assert.Same(importer, provider.GetRequiredService<IExcelImporter>());
        Assert.Same(exporter, provider.GetRequiredService<IExcelExporter>());
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(ICsvImporter))).Lifetime);
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(ICsvExporter))).Lifetime);
        Assert.All(services.Where(descriptor => descriptor.ServiceType == typeof(IExcelValidationRule)), descriptor =>
            Assert.Equal(ServiceLifetime.Singleton, descriptor.Lifetime));
    }

    private sealed class ReplacementImporter : IExcelImporter
    {
        public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(System.IO.Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request,
            System.Threading.CancellationToken cancellationToken = default)
            where TWorkbook : class, new() => throw new NotSupportedException();

        public System.Threading.Tasks.Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(
            System.IO.Stream source, ExcelWorkbookImportRequest<TWorkbook> request,
            System.Threading.CancellationToken cancellationToken = default)
            where TWorkbook : class, new() => throw new NotSupportedException();
    }

    private sealed class ReplacementExporter : IExcelExporter
    {
        public void Export(ExcelWorkbookExportRequest request, System.IO.Stream destination,
            System.Threading.CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public void ExportToFile(ExcelWorkbookExportRequest request, string path,
            System.Threading.CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public System.Threading.Tasks.Task ExportAsync(ExcelWorkbookExportRequest request,
            System.IO.Stream destination,
            System.Threading.CancellationToken cancellationToken = default) => throw new NotSupportedException();

        public System.Threading.Tasks.Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
            System.Threading.CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }
}
