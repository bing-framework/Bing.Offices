using Bing.Offices.Imports;

namespace Bing.Offices.Extensions;

/// <summary>无需声明 Workbook 根模型的单表导入入口；所有处理均委托现有 Workbook 管线。</summary>
public static class ExcelSingleSheetImportExtensions
{
    /// <summary>从调用方拥有的流导入一个 Sheet，调用完成后保持源流打开。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="source">可读取的 Excel 源流。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0，不跳过隐藏表。</param>
    /// <param name="configure">可选的 Sheet 配置，沿用表头、映射、校验和范围默认值。</param>
    /// <param name="resourceLimits">可选的导入资源上限；null 沿用 Provider 默认值。</param>
    /// <param name="cancellationToken">取消令牌；取消异常原样传播。</param>
    /// <returns>成功实体、原始错误和源行信息。</returns>
    public static ExcelSingleSheetImportResult<T> ImportSheet<T>(this IExcelImporter importer, Stream source,
        ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        return ToResult(importer.Import(source, CreateRequest(selector, configure, resourceLimits), cancellationToken));
    }

    /// <summary>使用现有异步 Workbook 管线导入一个 Sheet；仅外围 IO 异步，NPOI DOM 阶段仍同步。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="source">调用方拥有的可读流，调用后保持打开。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0。</param>
    /// <param name="configure">可选的 Sheet 配置。</param>
    /// <param name="resourceLimits">可选的资源上限。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含原始错误和源行的单表结果。</returns>
    public static async Task<ExcelSingleSheetImportResult<T>> ImportSheetAsync<T>(this IExcelImporter importer,
        Stream source, ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        var result = await importer.ImportAsync(source, CreateRequest(selector, configure, resourceLimits),
            cancellationToken).ConfigureAwait(false);
        return ToResult(result);
    }

    /// <summary>从文件导入一个 Sheet，使用既有文件扩展管理源文件生命周期。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="path">源文件路径。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0。</param>
    /// <param name="configure">可选的 Sheet 配置。</param>
    /// <param name="resourceLimits">可选的资源上限。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含原始错误和源行的单表结果。</returns>
    public static ExcelSingleSheetImportResult<T> ImportSheetFromFile<T>(this IExcelImporter importer,
        string path, ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        return ToResult(importer.ImportFromFile(path, CreateRequest(selector, configure, resourceLimits),
            cancellationToken));
    }

    /// <summary>从文件异步导入一个 Sheet，复用既有真实异步文件 IO。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="path">源文件路径。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0。</param>
    /// <param name="configure">可选的 Sheet 配置。</param>
    /// <param name="resourceLimits">可选的资源上限。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含原始错误和源行的单表结果。</returns>
    public static async Task<ExcelSingleSheetImportResult<T>> ImportSheetFromFileAsync<T>(this IExcelImporter importer,
        string path, ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        var result = await importer.ImportFromFileAsync(path, CreateRequest(selector, configure, resourceLimits),
            cancellationToken).ConfigureAwait(false);
        return ToResult(result);
    }

    /// <summary>从 Excel 字节数组导入一个 Sheet。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="content">Excel 文件内容。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0。</param>
    /// <param name="configure">可选的 Sheet 配置。</param>
    /// <param name="resourceLimits">可选的资源上限。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含原始错误和源行的单表结果。</returns>
    public static ExcelSingleSheetImportResult<T> ImportSheetFromBytes<T>(this IExcelImporter importer,
        byte[] content, ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        return ToResult(importer.ImportFromBytes(content, CreateRequest(selector, configure, resourceLimits),
            cancellationToken));
    }

    /// <summary>通过既有异步 Workbook 管线从 Excel 字节数组导入一个 Sheet。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="importer">执行 Workbook 管线的导入器。</param>
    /// <param name="content">Excel 文件内容。</param>
    /// <param name="selector">工作表选择器；null 表示物理索引 0。</param>
    /// <param name="configure">可选的 Sheet 配置。</param>
    /// <param name="resourceLimits">可选的资源上限。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含原始错误和源行的单表结果。</returns>
    public static async Task<ExcelSingleSheetImportResult<T>> ImportSheetFromBytesAsync<T>(this IExcelImporter importer,
        byte[] content, ExcelSheetSelector selector = null, Action<ExcelSheetImportBuilder<T>> configure = null,
        ExcelResourceLimits resourceLimits = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (importer == null)
            throw new ArgumentNullException(nameof(importer));
        var result = await importer.ImportFromBytesAsync(content, CreateRequest(selector, configure, resourceLimits),
            cancellationToken).ConfigureAwait(false);
        return ToResult(result);
    }

    /// <summary>创建唯一的单表请求，保持原有默认配置和可扩展的 Workbook 执行入口。</summary>
    private static ExcelWorkbookImportRequest<SingleSheetWorkbook<T>> CreateRequest<T>(ExcelSheetSelector selector,
        Action<ExcelSheetImportBuilder<T>> configure, ExcelResourceLimits limits) where T : class, new()
        => ExcelImport.Workbook<SingleSheetWorkbook<T>>(builder => builder.ResourceLimits(limits)
            .Sheet(selector ?? ExcelSheetSelector.ByIndex(0), root => root.Items, configure));

    /// <summary>保留全局错误，避免缺 Sheet 时因没有 Sheet 视图而丢失诊断。</summary>
    private static ExcelSingleSheetImportResult<T> ToResult<T>(ExcelWorkbookImportResult<SingleSheetWorkbook<T>> result)
        where T : class, new()
        => new(result.Workbook.Items, result.Errors,
            result.Sheets.Count == 0 ? Array.Empty<int>() : result.Sheets[0].SourceRows,
            result.ErrorsTruncated, result.MaxErrors);

    /// <summary>为 Workbook 管线提供已初始化的目标集合，不要求用户声明根模型。</summary>
    private sealed class SingleSheetWorkbook<T> where T : class, new()
    {
        /// <summary>创建用于本次请求的独立根模型。</summary>
        public SingleSheetWorkbook() { }

        /// <summary>获取接收成功实体的集合。</summary>
        public List<T> Items { get; } = new();
    }
}
