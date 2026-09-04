using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Mappings;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 服务扩展
/// </summary>
public static class ExcelNpoiServiceCollectionExtensions
{
    /// <summary>
    /// 注册 Bing.Offices 的 NPOI 操作服务。
    /// </summary>
    /// <param name="services">服务集合</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesNpoi(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        foreach (var rule in ExcelValidationRules.CreateDefault())
            services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IExcelValidationRule), rule.GetType()));
        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<IExcelImporter, NpoiExcelImporter>();
        services.TryAddTransient<IExcelExporter, NpoiExcelExporter>();
        return services;
    }
}
