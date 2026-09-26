using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.MiniExcel.Internals;
using Bing.Offices.Providers;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelApi = MiniExcelLibs.MiniExcel;

namespace Bing.Offices.Exports;

/// <summary>
/// 基于 MiniExcel 的 XLSX 导出器。
/// </summary>
/// <remarks>
/// 将行数据以延迟字典序列交给 MiniExcel，避免预先构建完整 DOM。
/// </remarks>
public sealed class MiniExcelExcelExporter : IExcelExporter, IExcelProviderFeatureDescriptor
{
    /// <inheritdoc />
    public string ProviderName => "MiniExcel";

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

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
        "模板编辑、Workbook 原生规则、宏、ODS、XLSB 和加密请求明确拒绝。"
    };
    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.StreamingWorkbookCreation;
    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;
    /// <summary>
    /// 用于创建默认映射计划的值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>
    /// 创建导入和导出映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 将导出流原子提交到文件的提交器。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;
    /// <summary>
    /// 向异常观察器分发导出异常的分发器。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;
    /// <summary>
    /// 按运行时行类型创建 MiniExcel 映射计划的构建器。
    /// </summary>
    private readonly MiniExcelMappingPlanBuilder _planBuilder;

    /// <summary>
    /// 初始化一个 <see cref="MiniExcelExcelExporter" /> 类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    /// <param name="fileExportCommitter">原子文件提交器。</param>
    public MiniExcelExcelExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null,
        IFileExportCommitter fileExportCommitter = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault(
            valueConverters: _valueConverters);
        _planBuilder = new MiniExcelMappingPlanBuilder(_mappingPlanFactory);
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        ValidateArguments(request, destination, cancellationToken);
        try
        {
            ValidateRequest(request);
            var workbook = CreateWorkbookRows(request, cancellationToken);
            MiniExcelApi.SaveAs(destination, workbook, true, "Sheet1", ExcelType.XLSX,
                CreateConfiguration(request));
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (NotSupportedException exception)
        {
            var translated = new BingOfficesUnsupportedFeatureException(
                "当前 Excel 导出功能不受 MiniExcel 支持。", exception, "MiniExcel",
                BingOfficesOperation.Export, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("Excel 导出失败。", exception, "MiniExcel",
                BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            DisposeTemplate(request);
        }
    }

    /// <inheritdoc />
    public async Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        ValidateArguments(request, destination, cancellationToken);
        try
        {
            ValidateRequest(request);
            var workbook = CreateWorkbookRows(request, cancellationToken);
            await MiniExcelApi.SaveAsAsync(destination, workbook, true, "Sheet1", ExcelType.XLSX,
                CreateConfiguration(request), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (NotSupportedException exception)
        {
            var translated = new BingOfficesUnsupportedFeatureException(
                "当前 Excel 导出功能不受 MiniExcel 支持。", exception, "MiniExcel",
                BingOfficesOperation.Export, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("Excel 异步导出失败。", exception, "MiniExcel",
                BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            DisposeTemplate(request);
        }
    }

    /// <inheritdoc />
    public void ExportToFile(ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        try
        {
            _fileExportCommitter.Commit(path,
                destination => Export(request, destination, cancellationToken), cancellationToken, "MiniExcel");
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
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        try
        {
            await _fileExportCommitter.CommitAsync(path,
                (destination, token) => ExportAsync(request, destination, token), cancellationToken,
                "MiniExcel").ConfigureAwait(false);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <summary>
    /// 创建按工作表名称组织的延迟行序列。
    /// </summary>
    /// <param name="request">包含工作表和映射配置的工作簿导出请求。</param>
    /// <param name="cancellationToken">用于取消行序列创建的令牌。</param>
    /// <returns>以工作表名称为键、以行序列为值的工作簿数据。</returns>
    private Dictionary<string, object> CreateWorkbookRows(ExcelWorkbookExportRequest request,
        CancellationToken cancellationToken)
    {
        var workbook = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateSheetName(sheet.Name);
            if (!names.Add(sheet.Name))
                throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheet.Name}");
            var mappingConfiguration = MiniExcelMappingPlanBuilder.MergeRequestDynamicColumns(
                sheet.MappingConfiguration, sheet.DynamicColumns);
            var plan = _planBuilder.Create(sheet.ItemType,
                sheet.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
                mappingConfiguration, MappingDirection.Export);
            ValidatePlanCapabilities(plan, sheet.Name);
            workbook.Add(sheet.Name, EnumerateRows(sheet, plan, cancellationToken));
        }
        return workbook;
    }

    /// <summary>
    /// 按映射计划枚举一个工作表的导出行。
    /// </summary>
    /// <param name="request">当前工作表导出请求。</param>
    /// <param name="plan">当前实体类型的导出映射计划。</param>
    /// <param name="cancellationToken">用于取消行枚举的令牌。</param>
    /// <returns>按工作表顺序生成的行字典序列。</returns>
    private IEnumerable<IDictionary<string, object>> EnumerateRows(ExcelSheetExportRequest request,
        IExcelMappingPlan plan, CancellationToken cancellationToken)
    {
        var fixedColumns = plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn
            && !IsDynamicContainerColumn(request.ItemType, column.Name)).ToArray();
        var dynamicColumns = plan.DynamicColumns.OrderBy(column => column.ColumnIndex ?? int.MaxValue)
            .ThenBy(column => column.Order).ToArray();
        if (fixedColumns.Length == 0 && dynamicColumns.Length == 0)
            throw new BingOfficesConfigurationException("MiniExcel 导出没有可写入的列。", stage: BingOfficesStage.Plan);

        // Sheet 计划在首行前固定反射元数据，避免每行重复 GetProperty 和动态键集合分配。
        var fixedBindings = fixedColumns.Select(column =>
        {
            var property = request.ItemType.GetProperty(column.Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (property == null)
                throw new BingOfficesConfigurationException($"属性不存在: {column.Name}",
                    stage: BingOfficesStage.Plan);
            return (Column: column, Property: property);
        }).ToArray();
        var knownDynamicKeys = new HashSet<string>(dynamicColumns.Select(column => column.Key),
            StringComparer.OrdinalIgnoreCase);

        var rowNumber = request.DataRowStartIndex + 1;
        foreach (var value in request.Data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (value == null)
                throw new ArgumentException($"Sheet {request.Name} 包含 null 数据项。", nameof(request));
            var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var propertyIndex = 0;
            foreach (var binding in fixedBindings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var raw = binding.Property.GetValue(value);
                row[binding.Column.Title] = MiniExcelValueAdapter.ConvertTo(raw, binding.Column, binding.Property,
                    request.Name, rowNumber, propertyIndex + 1, request.Culture ?? CultureInfo.InvariantCulture);
                propertyIndex++;
            }

            var dynamicValues = request.DynamicGetter?.Invoke(value);
            if (request.FailOnUnknownDynamicValues && dynamicValues != null
                && dynamicValues.Keys.Any(key => !knownDynamicKeys.Contains(key)))
            {
                var unknown = dynamicValues.Keys.First(key => !knownDynamicKeys.Contains(key));
                throw new BingOfficesConfigurationException($"动态列值未定义: {unknown}",
                    stage: BingOfficesStage.Plan);
            }
            foreach (var column in dynamicColumns)
            {
                object raw = null;
                dynamicValues?.TryGetValue(column.Key, out raw);
                row[column.Title] = MiniExcelValueAdapter.ConvertDynamicTo(raw, column,
                    request.Name, rowNumber, propertyIndex + 1,
                    request.Culture ?? CultureInfo.InvariantCulture);
                propertyIndex++;
            }
            yield return row;
            rowNumber++;
        }
    }

    /// <summary>
    /// 根据工作簿请求创建 MiniExcel Open XML 配置。
    /// </summary>
    /// <param name="request">用于读取区域性设置的工作簿导出请求。</param>
    /// <returns>启用空值单元格和共享字符串缓存的 Open XML 配置。</returns>
    private static OpenXmlConfiguration CreateConfiguration(ExcelWorkbookExportRequest request)
    {
        var culture = request.Sheets.FirstOrDefault()?.Culture;
        return new OpenXmlConfiguration
        {
            Culture = culture ?? CultureInfo.InvariantCulture,
            EnableWriteNullValueCell = true,
            EnableSharedStringCache = true
        };
    }

    /// <summary>
    /// 判断属性是否为动态列容器。
    /// </summary>
    /// <param name="itemType">行实体类型。</param>
    /// <param name="propertyName">待检查的属性名称。</param>
    /// <returns>属性存在且可赋值为字符串对象字典时返回 true，否则返回 false。</returns>
    private static bool IsDynamicContainerColumn(Type itemType, string propertyName)
    {
        var property = itemType.GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
        return property != null && typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType);
    }

    /// <summary>
    /// 校验导出请求、目标流和取消状态。
    /// </summary>
    /// <param name="request">待导出的工作簿请求。</param>
    /// <param name="destination">接收导出内容的目标流。</param>
    /// <param name="cancellationToken">用于取消导出的令牌。</param>
    private static void ValidateArguments(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 校验 MiniExcel 导出请求是否使用受支持的功能。
    /// </summary>
    /// <param name="request">待校验的工作簿导出请求。</param>
    private static void ValidateRequest(ExcelWorkbookExportRequest request)
    {
        if (request.Sheets.Any(sheet => sheet.Images.Count > 0 || sheet.DataValidations.Count > 0))
            throw new BingOfficesUnsupportedFeatureException("MiniExcel 不支持公共图片或原生数据校验导出。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);
        if (request.Format != ExcelFormat.Xlsx)
            throw new BingOfficesUnsupportedFeatureException("MiniExcel Provider 仅支持 XLSX。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
        if (request.MetadataSpecified)
            throw new BingOfficesUnsupportedFeatureException("MiniExcel Provider 暂不支持工作簿元数据。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
        if (request.Template != null)
            throw new BingOfficesUnsupportedFeatureException("MiniExcel Provider 暂不支持模板导出。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
        foreach (var sheet in request.Sheets)
        {
            if (sheet.Charts?.Count > 0 || sheet.HeaderRows?.Count > 0)
                throw new BingOfficesUnsupportedFeatureException("MiniExcel Provider 暂不支持图表或多级表头。",
                    provider: "MiniExcel", operation: BingOfficesOperation.Export,
                    stage: BingOfficesStage.Preflight);
            if (sheet.HeaderRowIndex != 0 || sheet.DataRowStartIndex != 1)
                throw new BingOfficesUnsupportedFeatureException("MiniExcel Provider 仅支持默认表头/数据行位置。",
                    provider: "MiniExcel", operation: BingOfficesOperation.Export,
                    stage: BingOfficesStage.Preflight);
            if (!string.IsNullOrWhiteSpace(sheet.TemplateRegion)
                || sheet.SheetStyle != null || sheet.HeaderStyle != null || sheet.BodyStyle != null
                || sheet.RowHeight != null
                || sheet.Hidden
                || (sheet.ColumnWidth != null && sheet.ColumnWidth.Mode != ExcelColumnWidthMode.None)
                || sheet.CommentConflictPolicy != ExcelCommentConflictPolicy.Preserve
                || sheet.TemplateCellOverwritePolicy != ExcelTemplateCellOverwritePolicy.PreserveTemplate)
                throw new BingOfficesUnsupportedFeatureException(
                    "MiniExcel Provider 暂不支持模板、样式、列宽、隐藏 Sheet 或批注相关导出选项。",
                    provider: "MiniExcel", operation: BingOfficesOperation.Export,
                    stage: BingOfficesStage.Preflight);
            if (sheet.DynamicColumns.Any(column => column != null
                    && (column.HeaderStyle != null || column.BodyStyle != null
                        || !string.IsNullOrWhiteSpace(column.NumberFormat)
                        || column.Placement != null || column.PhysicalColumnIndex.HasValue
                        || column.ImageMultiplicity != ExcelImageMultiplicityPolicy.First)))
                throw new BingOfficesUnsupportedFeatureException(
                    "MiniExcel Provider 暂不支持动态列样式、数字格式、物理定位或图片策略。",
                    provider: "MiniExcel", operation: BingOfficesOperation.Export,
                    stage: BingOfficesStage.Preflight);
        }
    }

    /// <summary>
    /// 校验映射计划是否超出 MiniExcel 的导出能力。
    /// </summary>
    /// <param name="plan">待校验的工作表映射计划。</param>
    /// <param name="sheetName">用于错误消息的工作表名称。</param>
    private static void ValidatePlanCapabilities(IExcelMappingPlan plan, string sheetName)
    {
        if (plan?.Style != null && (!string.IsNullOrWhiteSpace(plan.Style.HeaderStyleKey)
                || !string.IsNullOrWhiteSpace(plan.Style.BodyStyleKey)
                || !string.IsNullOrWhiteSpace(plan.Style.NumberFormat)))
            throw new BingOfficesUnsupportedFeatureException(
                $"Sheet {sheetName} 的映射样式或数字格式暂不受 MiniExcel 支持。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
        if (plan?.Layout != null && (plan.Layout.ColumnIndex.HasValue
                || !string.IsNullOrWhiteSpace(plan.Layout.PlacementKey)))
            throw new BingOfficesUnsupportedFeatureException(
                $"Sheet {sheetName} 的映射布局暂不受 MiniExcel 支持。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
        if (plan?.Columns?.Any(column => column.ImageMultiplicity != ExcelImageMultiplicityPolicy.First) == true
            || plan?.DynamicColumns?.Any(column => column.ImageMultiplicity != ExcelImageMultiplicityPolicy.First
                || !string.IsNullOrWhiteSpace(column.NumberFormat)
                || column.ColumnIndex.HasValue
                || !string.IsNullOrWhiteSpace(column.PlacementKey)) == true)
            throw new BingOfficesUnsupportedFeatureException(
                $"Sheet {sheetName} 的图片策略、动态列数字格式或布局暂不受 MiniExcel 支持。",
                provider: "MiniExcel", operation: BingOfficesOperation.Export,
                stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 校验工作表名称符合 XLSX 名称限制。
    /// </summary>
    /// <param name="name">待校验的工作表名称。</param>
    private static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31 || name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称无效: {name}", nameof(name));
    }

    /// <summary>
    /// 根据请求设置释放导出模板资源。
    /// </summary>
    /// <param name="request">可能包含模板流的工作簿导出请求。</param>
    private static void DisposeTemplate(ExcelWorkbookExportRequest request)
    {
        if (request?.Template != null && !request.LeaveTemplateOpen)
            request.Template.Dispose();
    }
}
