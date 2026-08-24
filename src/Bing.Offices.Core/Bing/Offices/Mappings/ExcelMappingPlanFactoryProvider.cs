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
    public static IExcelMappingPlanFactory CreateDefault(
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null) =>
        new ExcelMappingPlanFactory(valueConverters: valueConverters,
            validationRules: validationRules, namedValidationRules: namedValidationRules);

    /// <summary>
    /// 由 Core 注册默认映射计划工厂；已预注册的实现保持优先。
    /// </summary>
    public static IServiceCollection RegisterDefault(IServiceCollection services)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));
        services.TryAddSingleton<IExcelMappingPlanFactory>(provider =>
            CreateDefault(provider.GetServices<IExcelValueConverter>(),
                provider.GetServices<IExcelValidationRule>(),
                provider.GetServices<INamedExcelValidationRule>()));
        return services;
    }
}
