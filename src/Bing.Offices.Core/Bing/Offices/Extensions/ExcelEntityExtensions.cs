using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.Extensions;

/// <summary>
/// 提供实体和模板操作的统一扩展入口。
/// </summary>
public static class ExcelEntityExtensions
{
    /// <summary>
    /// 从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>实体导入结果。</returns>
    public static ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (importer is IExcelEntityImporter entityImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | RequiredLayoutCapabilities(layout),
                BingOfficesOperation.Import);
            return entityImporter.ImportEntity(source, layout, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import);
    }

    /// <summary>
    /// 从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>实体导入结果。</returns>
    public static ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityImportOptions options,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (importer is IExcelEntityResourceImporter resourceImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | RequiredLayoutCapabilities(layout),
                BingOfficesOperation.Import);
            return resourceImporter.ImportEntity(source, layout, options, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import,
            "Provider 未声明资源限制实体导入扩展接口。");
    }

    /// <summary>
    /// 异步从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>表示异步操作的任务，任务结果为实体导入结果。</returns>
    public static Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (importer is IExcelEntityImporter entityImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Async
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Import);
            return entityImporter.ImportEntityAsync(source, layout, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import);
    }

    /// <summary>
    /// 异步从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>表示异步操作的任务，任务结果为实体导入结果。</returns>
    public static Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityImportOptions options,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (importer is IExcelEntityResourceImporter resourceImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Async
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Import);
            return resourceImporter.ImportEntityAsync(source, layout, options, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import,
            "Provider 未声明资源限制实体导入扩展接口。");
    }

    /// <summary>
    /// 按模板约束从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与模板之间的布局定义。</param>
    /// <param name="template">模板导入选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>实体导入结果。</returns>
    public static ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (importer is IExcelEntityImporter entityImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Import);
            return entityImporter.ImportForTemplate(source, layout, template, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import);
    }

    /// <summary>
    /// 按模板约束从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="template">实体模板选项。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>实体导入结果。</returns>
    public static ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (importer is IExcelEntityResourceImporter resourceImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Import);
            return resourceImporter.ImportForTemplate(source, layout, template, options, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import,
            "Provider 未声明资源限制实体导入扩展接口。");
    }

    /// <summary>
    /// 异步按模板约束从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与模板之间的布局定义。</param>
    /// <param name="template">模板导入选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>表示异步操作的任务，任务结果为实体导入结果。</returns>
    public static Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (importer is IExcelEntityImporter entityImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | ExcelProviderCapabilities.Async | RequiredLayoutCapabilities(layout), BingOfficesOperation.Import);
            return entityImporter.ImportForTemplateAsync(source, layout, template, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import);
    }

    /// <summary>
    /// 异步按模板约束从工作簿流导入单个实体。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="importer">Excel 导入器。</param>
    /// <param name="source">包含工作簿内容的可读流。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="template">实体模板选项。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    /// <returns>表示异步操作的任务，任务结果为实体导入结果。</returns>
    public static Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(this IExcelImporter importer,
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (importer is IExcelEntityResourceImporter resourceImporter)
        {
            EnsureCapability(importer, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | ExcelProviderCapabilities.Async | RequiredLayoutCapabilities(layout),
                BingOfficesOperation.Import);
            return resourceImporter.ImportForTemplateAsync(source, layout, template, options, cancellationToken);
        }
        throw Unsupported(importer, BingOfficesOperation.Import,
            "Provider 未声明资源限制实体导入扩展接口。");
    }

    /// <summary>
    /// 将单个实体导出到工作簿流。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static void ExportEntity<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, Stream destination, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | RequiredLayoutCapabilities(layout),
                BingOfficesOperation.Export);
            entityExporter.ExportEntity(entity, layout, destination, cancellationToken);
            return;
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 异步将单个实体导出到工作簿流。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static Task ExportEntityAsync<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, Stream destination, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Async
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Export);
            return entityExporter.ExportEntityAsync(entity, layout, destination, cancellationToken);
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 以原子提交方式将单个实体导出到文件。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static void ExportEntityToFile<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, string path, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | RequiredLayoutCapabilities(layout),
                BingOfficesOperation.Export);
            entityExporter.ExportEntityToFile(entity, layout, path, cancellationToken);
            return;
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 异步以原子提交方式将单个实体导出到文件。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与工作簿之间的布局定义。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static Task ExportEntityToFileAsync<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, string path, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Async
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Export);
            return entityExporter.ExportEntityToFileAsync(entity, layout, path, cancellationToken);
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 使用模板将单个实体导出到工作簿流。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与模板之间的布局定义。</param>
    /// <param name="template">模板导出选项。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static void ExportForTemplate<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | RequiredLayoutCapabilities(layout), BingOfficesOperation.Export);
            entityExporter.ExportForTemplate(entity, layout, template, destination, cancellationToken);
            return;
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 异步使用模板将单个实体导出到工作簿流。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="exporter">Excel 导出器。</param>
    /// <param name="entity">要导出的实体。</param>
    /// <param name="layout">实体与模板之间的布局定义。</param>
    /// <param name="template">模板导出选项。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消操作的令牌。</param>
    public static Task ExportForTemplateAsync<TEntity>(this IExcelExporter exporter, TEntity entity,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (exporter == null)
            throw new ArgumentNullException(nameof(exporter));
        if (exporter is IExcelEntityExporter entityExporter)
        {
            EnsureCapability(exporter, ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Template
                | ExcelProviderCapabilities.Async | RequiredLayoutCapabilities(layout), BingOfficesOperation.Export);
            return entityExporter.ExportForTemplateAsync(entity, layout, template, destination, cancellationToken);
        }
        throw Unsupported(exporter, BingOfficesOperation.Export);
    }

    /// <summary>
    /// 创建 Provider 不支持实体操作时使用的统一异常。
    /// </summary>
    /// <param name="provider">被调用的 Provider 实例。</param>
    /// <param name="operation">当前执行的 Office 操作。</param>
    /// <param name="message">可选的具体不支持原因。</param>
    /// <returns>包含 Provider 和处理阶段上下文的不支持功能异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(object provider,
        BingOfficesOperation operation, string message = null)
    {
        var name = provider is IExcelProviderCapabilities capabilities
            ? capabilities.ProviderName : provider.GetType().Name;
        return new BingOfficesUnsupportedFeatureException(
            message ?? $"Provider {name} 不支持 Entity/Template 操作。", provider: name, operation: operation,
            stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 校验 Provider 是否声明了布局操作所需的全部能力。
    /// </summary>
    /// <param name="provider">被调用的 Provider 实例。</param>
    /// <param name="required">当前操作要求的能力组合。</param>
    /// <param name="operation">当前执行的 Office 操作。</param>
    private static void EnsureCapability(object provider, ExcelProviderCapabilities required,
        BingOfficesOperation operation)
    {
        if (!(provider is IExcelProviderCapabilities capabilities) || capabilities.Supports(required))
            return;
        throw new BingOfficesUnsupportedFeatureException(
            $"Provider {capabilities.ProviderName} 不支持能力: {required}。", provider: capabilities.ProviderName,
            operation: operation, stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 根据实体布局计算除实体能力外额外要求的 Provider 能力。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="layout">待执行的实体布局。</param>
    /// <returns>布局要求的附加 Provider 能力。</returns>
    private static ExcelProviderCapabilities RequiredLayoutCapabilities<TEntity>(ExcelEntityLayout<TEntity> layout)
        where TEntity : class, new() => layout != null && layout.Merges.Count > 0
            ? ExcelProviderCapabilities.Merge : ExcelProviderCapabilities.None;
}
