using Bing.Offices.Conversions;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.ExcelDataReader.Extensions;

/// <summary>
/// ExcelDataReader Provider 的依赖注入注册扩展。
/// </summary>
public static class ExcelDataReaderServiceCollectionExtensions
{
    /// <summary>
    /// 注册 ExcelDataReader 只读导入器。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>已完成注册的服务集合。</returns>
    public static IServiceCollection AddBingOfficesExcelDataReader(this IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        foreach (var rule in ExcelValidationRules.CreateDefault())
            services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IExcelValidationRule), rule.GetType()));
        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<ExcelDataReaderExcelImporter>(provider => new ExcelDataReaderExcelImporter(
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetService<IExcelMappingPlanFactory>()));
        services.TryAddTransient<IExcelImporter>(provider =>
            provider.GetRequiredService<ExcelDataReaderExcelImporter>());
        services.TryAddTransient<IExcelBatchImporter>(provider =>
            provider.GetRequiredService<ExcelDataReaderExcelImporter>());
        return services;
    }
}
