using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Starter.Provider;

/// <summary>
/// 最小 XLSX Provider 的服务注册入口。
/// </summary>
public static class StarterProviderServiceCollectionExtensions
{
    /// <summary>
    /// 注册基础 XLSX 导出能力。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>原始服务集合。</returns>
    public static IServiceCollection AddStarterExcelProvider(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));

        services.TryAddSingleton<IFileExportCommitter, DefaultFileExportCommitter>();
        services.TryAddTransient<StarterExcelExporter>(provider => new StarterExcelExporter(
            provider.GetRequiredService<IFileExportCommitter>(),
            provider.GetServices<IBingOfficesExceptionObserver>()));
        services.TryAddTransient<IExcelExporter>(provider => provider.GetRequiredService<StarterExcelExporter>());
        services.TryAddTransient<IExcelProviderFeatureDescriptor>(provider =>
            provider.GetRequiredService<StarterExcelExporter>());
        return services;
    }
}
