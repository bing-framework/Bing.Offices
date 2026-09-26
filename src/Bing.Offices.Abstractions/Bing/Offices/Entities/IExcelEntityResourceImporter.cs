using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 定义支持资源限制选项的实体导入扩展接口。
/// </summary>
/// <remarks>
/// 该接口独立于既有 <see cref="IExcelEntityImporter" />，避免新增成员破坏第三方 Provider 的既有实现。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityResourceImporter
{
    /// <summary>
    /// 使用资源限制从工作簿流导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>包含聚合对象、结构化错误和列表区域结果的导入结果。</returns>
    ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new();

    /// <summary>
    /// 异步使用资源限制从工作簿流导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>异步操作，结果包含聚合对象、结构化错误和列表区域结果。</returns>
    Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityImportOptions options,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 使用资源限制和模板布局导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">用于解析聚合对象的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>包含聚合对象、结构化错误和列表区域结果的导入结果。</returns>
    ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new();

    /// <summary>
    /// 异步使用资源限制和模板布局导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">用于解析聚合对象的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="options">实体导入资源限制选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>异步操作，结果包含聚合对象、结构化错误和列表区域结果。</returns>
    Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new();
}
