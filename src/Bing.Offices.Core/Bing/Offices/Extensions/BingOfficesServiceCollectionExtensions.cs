using Bing.Offices.Mappings;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Extensions;

/// <summary>注册不依赖 Excel Provider 的映射、校验和 CSV 服务。</summary>
public static class BingOfficesServiceCollectionExtensions
{
    /// <summary>注册 Core 默认服务和 CSV 导入导出器；重复注册幂等，保留调用方预注册的实现。</summary>
    /// <param name="services">依赖注入服务集合。</param>
    /// <returns>当前服务集合，可继续注册 NPOI 或自定义扩展。</returns>
    public static IServiceCollection AddBingOfficesCore(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        foreach (var rule in ExcelValidationRules.CreateDefault())
            services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IExcelValidationRule), rule.GetType()));
        return ExcelMappingPlanFactoryProvider.RegisterDefault(services);
    }
}
