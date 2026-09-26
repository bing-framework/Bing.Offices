using Bing.Offices.Conversions;
using Bing.Offices.Formula;
using Bing.Offices.Rendering;
using Bing.Offices.Providers;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.AsposeCells;

/// <summary>
/// Aspose.Cells Provider 的依赖注入扩展。
/// </summary>
public static class Extensions
{
    /// <summary>
    /// 注册渲染、公式、格式转换和页面图片能力。
    /// </summary>
    /// <param name="services">服务集合。</param>
    /// <param name="configure">可选宿主配置。</param>
    /// <returns>原服务集合。</returns>
    public static IServiceCollection AddBingOfficesAsposeCells(this IServiceCollection services,
        Action<AsposeCellsProviderOptions> configure = null)
    {
        if (services == null) throw new ArgumentNullException(nameof(services));
        var options = new AsposeCellsProviderOptions();
        configure?.Invoke(options);
        services.AddSingleton(options.Clone());
        services.TryAddSingleton<AsposeCellsEngine>();
        services.TryAddTransient<IExcelDocumentRenderer, AsposeCellsEngine>();
        services.TryAddTransient<IExcelPageRenderer, AsposeCellsEngine>();
        services.TryAddTransient<IExcelDocumentConverter, AsposeCellsEngine>();
        services.TryAddTransient<IExcelFormulaProcessor, AsposeCellsEngine>();
        services.TryAddTransient<IExcelProviderFeatureDescriptor, AsposeCellsEngine>();
        return services;
    }
}
