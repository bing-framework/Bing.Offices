using System;
using System.Collections.Generic;
using Bing.Offices.Conversions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Mappings;

/// <summary>
/// 为上层 Provider 提供默认 Plan 工厂的创建入口。
/// </summary>
public static class ExcelMappingPlanFactoryProvider
{
    /// <summary>创建默认的 Plan 工厂。</summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="profileRegistry">可选的方向 Profile 注册表。</param>
    /// <param name="modelAliases">可选的模型别名注册表。</param>
    public static IExcelMappingPlanFactory CreateDefault(
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        Configurations.IMappingProfileRegistry profileRegistry = null,
        Configurations.ExcelModelAliasRegistry modelAliases = null) =>
        new ExcelMappingPlanFactory(valueConverters, validationRules, namedValidationRules,
            profileRegistry: profileRegistry, modelAliases: modelAliases);

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
                provider.GetService<Configurations.IMappingProfileRegistry>(),
                provider.GetService<Configurations.ExcelModelAliasRegistry>()));
        return services;
    }
}
