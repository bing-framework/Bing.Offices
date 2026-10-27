using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Conversions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 服务扩展
/// </summary>
public static class ExcelNpoiServiceCollectionExtensions
{
    /// <summary>显式注册低内存 XLSX 服务，不替换既有完整工作簿服务或用户注册。</summary>
    /// <param name="services">服务集合。</param>
    /// <returns>原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesNpoiStreaming(this IServiceCollection services)
    {
        if (services == null) throw new ArgumentNullException(nameof(services));
        services.AddBingOfficesCore();
        services.TryAddTransient<IExcelStreamingExporter>(provider => new NpoiExcelStreamingExporter(
            provider.GetServices<IExcelValueConverter>(), provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(), provider.GetRequiredService<IFileExportCommitter>()));
        services.TryAddTransient<IExcelStreamingImporter>(provider => new NpoiExcelStreamingImporter(
            provider.GetServices<IExcelValidationRule>(), provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<INamedExcelValidationRule>(), provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>()));
        return services;
    }

    /// <summary>
    /// 注册 Bing.Offices 的 NPOI 操作服务。
    /// </summary>
    /// <param name="services">服务集合</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesNpoi(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        services.AddBingOfficesCore();
        services.TryAddTransient<IExcelImporter>(provider => new NpoiExcelImporter(
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>()));
        services.TryAddTransient<IExcelExporter>(provider => new NpoiExcelExporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(),
            provider.GetRequiredService<IFileExportCommitter>()));
        return services;
    }
}
