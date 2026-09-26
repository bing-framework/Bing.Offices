using Bing.Offices.Mappings;
using Bing.Offices.Validations;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Conversions;
using Bing.Offices.Formula;
using Bing.Offices.ClosedXml.Formula;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.ClosedXml.Extensions;

/// <summary>
/// ClosedXML Provider 的依赖注入注册扩展。
/// </summary>
public static class ExcelClosedXmlServiceCollectionExtensions
{
    /// <summary>
    /// 注册 ClosedXML Provider 及其默认依赖。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesClosedXml(this IServiceCollection services)
        => AddBingOfficesClosedXml(services, null);

    /// <summary>
    /// 注册 ClosedXML Provider，并配置 Workbook DOM 准入限制。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <param name="configure">Provider 选项配置委托；为空时使用默认选项。</param>
    /// <returns>已完成注册的原始服务集合。</returns>
    public static IServiceCollection AddBingOfficesClosedXml(this IServiceCollection services,
        Action<ClosedXmlProviderOptions> configure)
    {
        if (services == null)
            throw new ArgumentNullException(nameof(services));

        var options = new ClosedXmlProviderOptions();
        configure?.Invoke(options);
        options.Validate();
        services.TryAddSingleton(options);
        services.TryAddSingleton<IClosedXmlWorkbookAdmission>(provider =>
            new ClosedXmlWorkbookAdmission(provider.GetRequiredService<ClosedXmlProviderOptions>()));

        foreach (var rule in ExcelValidationRules.CreateDefault())
            services.TryAddEnumerable(ServiceDescriptor.Singleton(typeof(IExcelValidationRule), rule.GetType()));
        ExcelMappingPlanFactoryProvider.RegisterDefault(services);
        services.TryAddTransient<IExcelImporter>(provider => new ClosedXmlExcelImporter(
            provider.GetServices<IExcelValidationRule>(),
            provider.GetServices<IExcelValueConverter>(),
            provider.GetServices<INamedExcelValidationRule>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(),
            provider.GetRequiredService<IClosedXmlWorkbookAdmission>()));
        services.TryAddTransient<IExcelExporter>(provider => new ClosedXmlExcelExporter(
            provider.GetServices<IExcelValueConverter>(),
            provider.GetService<IExcelMappingPlanFactory>(),
            provider.GetServices<IBingOfficesExceptionObserver>(),
            provider.GetRequiredService<IFileExportCommitter>(),
            provider.GetRequiredService<IClosedXmlWorkbookAdmission>()));
        services.TryAddTransient<IExcelFormulaProcessor, ClosedXmlExcelFormulaProcessor>();
        services.TryAddTransient<IExcelEntityImporter>(provider =>
            (IExcelEntityImporter)provider.GetRequiredService<IExcelImporter>());
        services.TryAddTransient<IExcelEntityExporter>(provider =>
            (IExcelEntityExporter)provider.GetRequiredService<IExcelExporter>());
        return services;
    }
}
