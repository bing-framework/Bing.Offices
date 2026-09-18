using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.MiniExcel.Extensions;

/// <summary>
/// MiniExcel Provider 的依赖注入注册扩展。
/// </summary>
public static class ExcelMiniExcelServiceCollectionExtensions
{
    /// <summary>
    /// 注册 Bing.Offices 的 MiniExcel 操作服务。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesMiniExcel(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        foreach (var rule in ExcelValidationRules.CreateDefault())
            services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IExcelValidationRule), rule.GetType()));
        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<IExcelImporter>(provider => new MiniExcelExcelImporter(
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>()));
        services.TryAddTransient<IExcelExporter>(provider => new MiniExcelExcelExporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(),
            provider.GetRequiredService<IFileExportCommitter>()));
        return services;
    }
}
