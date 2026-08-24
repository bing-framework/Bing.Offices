using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
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
    /// 注册Npoi操作
    /// </summary>
    /// <param name="services">服务集合</param>
    public static void AddNpoi(this IServiceCollection services)
    {
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, RequiredExcelValidationRule>());
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, RegexExcelValidationRule>());
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, RangeExcelValidationRule>());
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, MaxLengthExcelValidationRule>());
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, DateTimeExcelValidationRule>());
        services.TryAddEnumerable(ServiceDescriptor.Singleton<IExcelValidationRule, DuplicationExcelValidationRule>());
        services.TryAddSingleton<IExcelMappingConfigurationLoader, DefaultExcelMappingConfigurationLoader>();
        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<IExcelImporter, NpoiExcelImporter>();
        services.TryAddTransient<IExcelExporter, NpoiExcelExporter>();
        services.TryAddTransient<ICsvImporter, CsvEntityImporter>();
        services.TryAddTransient<ICsvExporter, CsvEntityExporter>();
    }
}
