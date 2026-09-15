using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Csv;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Mappings;

/// <summary>
/// 为上层 Provider 提供默认 Plan 工厂的创建入口。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public static class ExcelMappingPlanFactoryProvider
{
    /// <summary>创建默认的 Plan 工厂。</summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="profileRegistry">可选的方向 Profile 注册表。</param>
    /// <param name="modelAliases">可选的模型别名注册表。</param>
    /// <param name="cacheCapacity">映射计划缓存的最大条目数。</param>
    /// <returns>使用给定扩展点和缓存容量创建的映射计划工厂。</returns>
    public static IExcelMappingPlanFactory CreateDefault(
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        Configurations.IMappingProfileResolver profileRegistry = null,
        Configurations.ExcelModelAliasRegistry modelAliases = null,
        int cacheCapacity = 256) =>
        new ExcelMappingPlanFactory(valueConverters, validationRules, namedValidationRules,
            cacheCapacity, profileRegistry, modelAliases);

    /// <summary>
    /// 由 Core 注册默认映射计划工厂；已预注册的实现保持优先。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>已注册默认 Core 服务的同一服务集合。</returns>
    public static IServiceCollection RegisterDefault(IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        services.TryAddSingleton<IExcelMappingPlanFactory>(provider =>
            CreateDefault(
                provider.GetServices<IExcelValueConverter>(),
                provider.GetServices<IExcelValidationRule>(),
                provider.GetServices<INamedExcelValidationRule>(),
                provider.GetService<Configurations.IMappingProfileResolver>(),
                provider.GetService<Configurations.ExcelModelAliasRegistry>(),
                256));
        services.TryAddSingleton<IExcelMappingConfigurationLoader, DefaultExcelMappingConfigurationLoader>();
        services.TryAddSingleton<IFileExportCommitter, DefaultFileExportCommitter>();
        services.TryAddTransient<ICsvImporter>(provider => new CsvEntityImporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetRequiredService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>()));
        services.TryAddTransient<ICsvExporter>(provider => new CsvEntityExporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetRequiredService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(),
            provider.GetRequiredService<IFileExportCommitter>()));
        return services;
    }
}
