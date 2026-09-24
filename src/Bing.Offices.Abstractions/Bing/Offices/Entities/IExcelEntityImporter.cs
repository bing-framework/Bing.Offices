using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 定义支持固定单元格、合并区域、列表区域和父子关系的实体导入接口。
/// </summary>
/// <remarks>
/// 模板导入方法使用 source 作为实际数据工作簿，template 作为结构预检工作簿。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityImporter
{
    /// <summary>
    /// 从工作簿流导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>包含聚合对象、结构化错误和列表区域结果的导入结果。</returns>
    ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 异步从工作簿流导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>表示异步导入的任务，任务结果包含聚合对象和结构化错误。</returns>
    Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
        where TEntity : class, new();

    /// <summary>
    /// 按既有模板布局导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">用于解析聚合对象的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>包含聚合对象、结构化错误和列表区域结果的导入结果。</returns>
    ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        ExcelEntityTemplateOptions template, CancellationToken cancellationToken = default)
        where TEntity : class, new();

    /// <summary>
    /// 异步按既有模板布局导入单个聚合对象。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">用于解析聚合对象的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="cancellationToken">用于取消导入操作的令牌。</param>
    /// <returns>表示异步模板导入的任务，任务结果包含聚合对象和结构化错误。</returns>
    Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        CancellationToken cancellationToken = default) where TEntity : class, new();
}
