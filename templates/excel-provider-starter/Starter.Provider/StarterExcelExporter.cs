using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using ClosedXML.Excel;

namespace Starter.Provider;

/// <summary>
/// 使用独立引擎写出基础 XLSX 工作簿的 Provider 示例。
/// </summary>
/// <remarks>
/// 仅支持公开属性的基础标量列和多个工作表；序列化先在内存中完成，再写入调用方目标。
/// 生产 Provider 应按自身资源预算替换暂存策略，并为新增能力补充职责测试。
/// </remarks>
public sealed class StarterExcelExporter : IExcelExporter, IExcelProviderFeatureDescriptor
{
    /// <summary>
    /// 单个工作表允许的最大数据行数。
    /// </summary>
    private const int MaxRowsPerSheet = 10000;

    /// <summary>
    /// 单个工作簿允许的最大序列化字节数。
    /// </summary>
    private const int MaxOutputBytes = 16 * 1024 * 1024;

    /// <summary>
    /// 安全提交目标文件的服务。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;

    /// <summary>
    /// 发送结构化异常的观察器。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个 <see cref="StarterExcelExporter"/> 类型的实例。
    /// </summary>
    /// <param name="fileExportCommitter">文件提交服务；为空时使用默认实现。</param>
    /// <param name="exceptionObservers">接收结构化异常的观察器。</param>
    public StarterExcelExporter(IFileExportCommitter fileExportCommitter = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public string ProviderName => "StarterClosedXml";

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List |
        ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.None;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = Array.Empty<ExcelFormat>();

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> WriteFormats { get; } = new[] { ExcelFormat.Xlsx };

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookImport => false;

    /// <inheritdoc />
    public bool SupportsBatchImport => false;

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookExport => true;

    /// <inheritdoc />
    public bool SupportsTrueAsyncIo => true;

    /// <inheritdoc />
    public IReadOnlyList<string> Limitations { get; } = new[]
    {
        "仅写入基础 XLSX 标量列表和多个工作表，不提供导入、Entity、模板或高级布局。",
        "每个工作表最多 10,000 行，序列化工作簿最多 16 MiB；输出前需要内存暂存。"
    };

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        ValidateDestination(destination);
        try
        {
            using var staged = Build(request, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            staged.CopyTo(destination);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException) { throw; }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw ExportFailure(exception);
        }
    }

    /// <inheritdoc />
    public async Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        ValidateDestination(destination);
        try
        {
            using var staged = Build(request, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            var bytes = staged.GetBuffer();
            for (var offset = 0; offset < staged.Length; offset += 81920)
            {
                var count = (int)Math.Min(81920, staged.Length - offset);
                await destination.WriteAsync(bytes.AsMemory(offset, count), cancellationToken)
                    .ConfigureAwait(false);
            }
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException) { throw; }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw ExportFailure(exception);
        }
    }

    /// <inheritdoc />
    public void ExportToFile(ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default)
    {
        ValidatePath(path);
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            _fileExportCommitter.Commit(path, destination => Export(request, destination, cancellationToken),
                cancellationToken, "XLSX");
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <inheritdoc />
    public async Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default)
    {
        ValidatePath(path);
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            await _fileExportCommitter.CommitAsync(path,
                (destination, token) => ExportAsync(request, destination, token),
                cancellationToken, "XLSX").ConfigureAwait(false);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <summary>
    /// 构建并暂存已通过预检的工作簿。
    /// </summary>
    /// <param name="request">导出请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>位置已重置的暂存流。</returns>
    private static MemoryStream Build(ExcelWorkbookExportRequest request, CancellationToken cancellationToken)
    {
        var columnsBySheet = ValidateRequest(request);
        cancellationToken.ThrowIfCancellationRequested();
        using var workbook = new XLWorkbook();
        for (var sheetIndex = 0; sheetIndex < request.Sheets.Count; sheetIndex++)
        {
            var sheet = request.Sheets[sheetIndex];
            var columns = columnsBySheet[sheetIndex];
            var worksheet = workbook.Worksheets.Add(sheet.Name);
            for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                worksheet.Cell(1, columnIndex + 1).Value = columns[columnIndex].Name;

            var rowIndex = 0;
            foreach (var item in sheet.Data)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (++rowIndex > MaxRowsPerSheet)
                    throw new BingOfficesResourceLimitException("工作表数据行数超过 10,000。",
                        provider: "StarterClosedXml", operation: BingOfficesOperation.Export,
                        stage: BingOfficesStage.Write);
                if (item == null || !sheet.ItemType.IsInstanceOfType(item))
                    throw new BingOfficesExportException("数据项类型与工作表模型不一致。",
                        provider: "StarterClosedXml", sheetName: sheet.Name, rowIndex: rowIndex + 1);
                for (var columnIndex = 0; columnIndex < columns.Length; columnIndex++)
                    WriteValue(worksheet.Cell(rowIndex + 1, columnIndex + 1), columns[columnIndex].GetValue(item));
            }
        }

        var staged = new MemoryStream();
        try
        {
            workbook.SaveAs(staged);
            if (staged.Length > MaxOutputBytes)
                throw new BingOfficesResourceLimitException("工作簿输出超过 16 MiB。",
                    provider: "StarterClosedXml", operation: BingOfficesOperation.Export,
                    stage: BingOfficesStage.Serialize);
            staged.Position = 0;
            return staged;
        }
        catch
        {
            staged.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 校验请求并规划公开标量列。
    /// </summary>
    /// <param name="request">导出请求。</param>
    /// <returns>按工作表顺序排列的属性列。</returns>
    private static IReadOnlyList<PropertyInfo[]> ValidateRequest(ExcelWorkbookExportRequest request)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (request.Format != ExcelFormat.Xlsx)
            throw Unsupported("仅支持 XLSX 格式。");
        if (request.Template != null || request.MetadataSpecified)
            throw Unsupported("不支持模板或工作簿元数据。");
        if (request.Sheets.Count == 0)
            throw Unsupported("至少需要一个工作表。");

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<PropertyInfo[]>();
        foreach (var sheet in request.Sheets)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(sheet.Name) || !names.Add(sheet.Name))
                throw Unsupported("工作表名称不能为空或重复。");
            if (sheet.Data == null || sheet.ItemType == null)
                throw Unsupported("工作表缺少数据或数据类型。");
            if (sheet.HeaderRowIndex != 0 || sheet.DataRowStartIndex != 1 || sheet.Hidden
                || sheet.DynamicColumns.Count != 0 || sheet.DynamicGetter != null
                || sheet.Charts.Count != 0 || sheet.HeaderRows.Count != 0
                || sheet.Tables.Count != 0 || sheet.AutoFilters.Count != 0
                || sheet.ConditionalFormats.Count != 0 || sheet.NamedRanges.Count != 0
                || sheet.Images.Count != 0 || sheet.DataValidations.Count != 0
                || sheet.FreezePane != null || sheet.PrintLayout != null
                || sheet.ColumnWidth != null || sheet.RowHeight != null
                || sheet.SheetStyle != null || sheet.HeaderStyle != null || sheet.BodyStyle != null
                || HasMapping(sheet.MappingConfiguration) || sheet.MappingDocument != null
                || !Equals(sheet.Culture, CultureInfo.InvariantCulture)
                || !string.IsNullOrEmpty(sheet.TemplateRegion))
                throw Unsupported($"工作表 {sheet.Name} 包含此模板尚未实现的布局或映射功能。");

            var columns = sheet.ItemType.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(property => property.GetMethod?.IsPublic == true
                    && property.GetIndexParameters().Length == 0)
                .OrderBy(property => property.Name, StringComparer.Ordinal).ToArray();
            if (columns.Length == 0 || columns.Length > 16384 || columns.Any(property =>
                    !IsScalar(property.PropertyType) || property.GetCustomAttributes(true).Any(attribute =>
                        attribute.GetType().Namespace?.StartsWith("Bing.Offices", StringComparison.Ordinal) == true)))
                throw Unsupported($"工作表 {sheet.Name} 只支持未标注映射特性的公开标量属性。");
            result.Add(columns);
        }
        return result;
    }

    /// <summary>
    /// 判断属性类型是否可直接写入单元格。
    /// </summary>
    /// <param name="type">属性类型。</param>
    /// <returns>支持时返回 true。</returns>
    private static bool IsScalar(Type type)
    {
        type = Nullable.GetUnderlyingType(type) ?? type;
        return type == typeof(string) || type == typeof(bool) || type == typeof(DateTime)
            || type == typeof(byte) || type == typeof(short) || type == typeof(int)
            || type == typeof(long) || type == typeof(float) || type == typeof(double)
            || type == typeof(decimal);
    }

    /// <summary>
    /// 判断请求级映射是否包含需要执行的配置。
    /// </summary>
    /// <param name="mapping">构建器生成的请求级映射。</param>
    /// <returns>配置了映射行为时返回 true。</returns>
    private static bool HasMapping(ExcelMappingConfiguration mapping) => mapping != null &&
        (!string.IsNullOrWhiteSpace(mapping.Profile) || !string.IsNullOrWhiteSpace(mapping.ModelAlias)
            || mapping.ClearDynamicColumns || mapping.DynamicColumnKeysToRemove.Count != 0
            || mapping.DynamicColumnMergeMode != null || mapping.ResetStyle || mapping.ResetLayout
            || mapping.Columns.Count != 0 || mapping.DynamicColumns.Count != 0
            || mapping.Style != null || mapping.Layout != null);

    /// <summary>
    /// 写入已验证的标量值。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="value">属性值。</param>
    private static void WriteValue(IXLCell cell, object value)
    {
        switch (value)
        {
            case null: break;
            case string text: cell.Value = text; break;
            case bool boolean: cell.Value = boolean; break;
            case DateTime date: cell.Value = date; break;
            default: cell.Value = Convert.ToDouble(value, CultureInfo.InvariantCulture); break;
        }
    }

    /// <summary>
    /// 校验调用方目标流。
    /// </summary>
    /// <param name="destination">目标流。</param>
    private static void ValidateDestination(Stream destination)
    {
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
    }

    /// <summary>
    /// 校验目标路径。
    /// </summary>
    /// <param name="path">目标路径。</param>
    private static void ValidatePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
    }

    /// <summary>
    /// 创建结构化不支持功能异常。
    /// </summary>
    /// <param name="message">限制原因。</param>
    /// <returns>预检阶段异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message) =>
        new BingOfficesUnsupportedFeatureException(message, provider: "StarterClosedXml",
            operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 包装并通知导出异常。
    /// </summary>
    /// <param name="exception">底层异常。</param>
    /// <returns>结构化导出异常。</returns>
    private BingOfficesExportException ExportFailure(Exception exception)
    {
        var result = new BingOfficesExportException("基础 XLSX 导出失败。", exception,
            provider: ProviderName, stage: BingOfficesStage.Write);
        _exceptionDispatcher.Observe(result);
        return result;
    }
}
