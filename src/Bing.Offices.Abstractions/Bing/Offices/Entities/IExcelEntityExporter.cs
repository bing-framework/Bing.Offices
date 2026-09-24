using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 定义支持固定单元格、合并区域、列表区域和父子关系的实体导出接口。
/// </summary>
/// <remarks>
/// 模板导出会保留模板中未映射的工作表内容、样式和资源，并只写入布局声明的区域。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityExporter
{
    /// <summary>
    /// 将单个聚合对象写入目标流。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消导出操作的令牌。</param>
    void ExportEntity<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 异步将单个聚合对象写入目标流。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="destination">接收工作簿内容的可写流。</param>
    /// <param name="cancellationToken">用于取消导出操作的令牌。</param>
    Task ExportEntityAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 以原子提交方式将单个聚合对象写入文件。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="cancellationToken">用于取消导出和文件提交的令牌。</param>
    void ExportEntityToFile<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, string path,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 异步以原子提交方式将单个聚合对象写入文件。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">固定单元格、合并区域和列表区域的布局。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="cancellationToken">用于取消导出和文件提交的令牌。</param>
    Task ExportEntityToFileAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, string path,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 使用既有模板布局将单个聚合对象写入目标流。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">用于填充模板的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="destination">接收填充后工作簿的可写流。</param>
    /// <param name="cancellationToken">用于取消导出操作的令牌。</param>
    void ExportForTemplate<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        ExcelEntityTemplateOptions template, Stream destination, CancellationToken cancellationToken = default)
        where TEntity : class, new();

    /// <summary>
    /// 异步使用既有模板布局将单个聚合对象写入目标流。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="entity">待导出的聚合对象。</param>
    /// <param name="layout">用于填充模板的布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="destination">接收填充后工作簿的可写流。</param>
    /// <param name="cancellationToken">用于取消导出操作的令牌。</param>
    Task ExportForTemplateAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        ExcelEntityTemplateOptions template, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new();
}
