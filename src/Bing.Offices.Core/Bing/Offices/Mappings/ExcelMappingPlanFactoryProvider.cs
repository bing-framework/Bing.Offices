using System;
using System.Collections.Generic;
using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Csv;
using Bing.Offices.Configurations;
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
        services.TryAddTransient<ICsvImporter>(provider => new CsvEntityImporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetRequiredService<IExcelMappingPlanFactory>()));
        services.TryAddTransient<ICsvExporter>(provider => new CsvEntityExporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetRequiredService<IExcelMappingPlanFactory>()));
        return services;
    }
}
