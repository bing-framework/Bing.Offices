using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.SpreadCheetah.Extensions;

/// <summary>
/// SpreadCheetah 流式导出 Provider 的依赖注入扩展。
/// </summary>
public static class ExcelSpreadCheetahServiceCollectionExtensions
{
    /// <summary>
    /// 注册只创建新 XLSX 的流式导出能力。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesSpreadCheetah(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));

        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<IExcelStreamingExporter>(provider =>
            new SpreadCheetahStreamingExcelExporter(
                provider.GetRequiredService<IExcelMappingPlanFactory>(),
                provider.GetRequiredService<IFileExportCommitter>()));
        services.TryAddTransient<IExcelProviderFeatureDescriptor>(provider =>
            (IExcelProviderFeatureDescriptor)provider.GetRequiredService<IExcelStreamingExporter>());
        return services;
    }
}
