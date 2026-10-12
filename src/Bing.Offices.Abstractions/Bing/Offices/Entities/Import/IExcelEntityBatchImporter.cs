using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Bing.Offices.Entities;

/// <summary>
/// 定义按实体列表区域分批导入聚合对象的 Provider SPI。
/// </summary>
/// <remarks>
/// 批次 Provider 先完成实体布局导入，再按列表区域投影批次；不代表前向解析或低内存语义。
/// </remarks>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelEntityBatchImporter
{
    /// <summary>
    /// 按实体列表区域同步交付导入批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="onBatch">每个区域批次的同步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的批次导入摘要。</returns>
    ExcelEntityBatchImportSummary<TEntity> ImportEntityBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按实体列表区域使用指定选项同步交付导入批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="onBatch">每个区域批次的同步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的批次导入摘要。</returns>
    ExcelEntityBatchImportSummary<TEntity> ImportEntityBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按实体列表区域异步交付导入批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="onBatch">每个区域批次的异步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的异步批次导入摘要。</returns>
    Task<ExcelEntityBatchImportSummary<TEntity>> ImportEntityBatchesAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按实体列表区域使用指定选项异步交付导入批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="onBatch">每个区域批次的异步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的异步批次导入摘要。</returns>
    Task<ExcelEntityBatchImportSummary<TEntity>> ImportEntityBatchesAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按既有模板布局同步交付实体列表区域批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="onBatch">每个区域批次的同步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的批次导入摘要。</returns>
    ExcelEntityBatchImportSummary<TEntity> ImportForTemplateBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按既有模板布局使用指定选项同步交付实体列表区域批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="onBatch">每个区域批次的同步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的批次导入摘要。</returns>
    ExcelEntityBatchImportSummary<TEntity> ImportForTemplateBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityBatchImportOptions options, Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按既有模板布局异步交付实体列表区域批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="onBatch">每个区域批次的异步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的异步批次导入摘要。</returns>
    Task<ExcelEntityBatchImportSummary<TEntity>> ImportForTemplateBatchesAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();

    /// <summary>
    /// 按既有模板布局使用指定选项异步交付实体列表区域批次。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="source">包含待导入工作簿的可读流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="template">提供模板流及其所有权设置的选项。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="onBatch">每个区域批次的异步回调。</param>
    /// <param name="cancellationToken">用于取消导入和批次交付的令牌。</param>
    /// <returns>包含完整实体结果和各区域摘要的异步批次导入摘要。</returns>
    Task<ExcelEntityBatchImportSummary<TEntity>> ImportForTemplateBatchesAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new();
}
