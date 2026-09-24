using System.Collections;
using System.Globalization;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Exports;

/// <summary>
/// 基于 ClosedXML 的 XLSX 工作簿导出器。
/// </summary>
public sealed class ClosedXmlExcelExporter : IExcelExporter, IExcelEntityExporter, IExcelProviderCapabilities
{
    /// <summary>
    /// ClosedXML Provider 名称。
    /// </summary>
    private const string Provider = "ClosedXML";

    /// <summary>
    /// 公共映射计划工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;

    /// <summary>
    /// 原子文件提交器。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;

    /// <summary>
    /// Bing.Offices 异常观察分发器。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// ClosedXML 映射计划构建器。
    /// </summary>
    private readonly ClosedXmlMappingPlanBuilder _planBuilder;

    /// <summary>
    /// 实体布局执行器。
    /// </summary>
    private readonly ClosedXmlEntityLayoutExecutor _entityExecutor;

    /// <summary>
    /// Workbook DOM 准入器。
    /// </summary>
    private readonly IClosedXmlWorkbookAdmission _admission;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelExporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 使用进程级默认 Workbook DOM 准入器。
    /// </remarks>
    /// <param name="valueConverters">可选的值转换器集合。</param>
    /// <param name="mappingPlanFactory">可选的公共映射计划工厂。</param>
    /// <param name="exceptionObservers">可选的异常观察器集合。</param>
    /// <param name="fileExportCommitter">可选的原子文件提交器。</param>
    public ClosedXmlExcelExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null,
        IFileExportCommitter fileExportCommitter = null)
        : this(valueConverters, mappingPlanFactory, exceptionObservers, fileExportCommitter,
            ClosedXmlWorkbookAdmission.SharedDefault)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelExporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 使用调用方提供的 Workbook DOM 准入器。
    /// </remarks>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">公共映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    /// <param name="fileExportCommitter">原子文件提交器。</param>
    /// <param name="admission">Workbook DOM 准入器。</param>
    internal ClosedXmlExcelExporter(IEnumerable<IExcelValueConverter> valueConverters,
        IExcelMappingPlanFactory mappingPlanFactory,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers,
        IFileExportCommitter fileExportCommitter,
        IClosedXmlWorkbookAdmission admission)
    {
        var converters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault(
            valueConverters: converters);
        _planBuilder = new ClosedXmlMappingPlanBuilder(_mappingPlanFactory);
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
        _entityExecutor = new ClosedXmlEntityLayoutExecutor(_mappingPlanFactory);
        _admission = admission ?? throw new ArgumentNullException(nameof(admission));
    }

    /// <inheritdoc />
    public string ProviderName => Provider;

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Entity
        | ExcelProviderCapabilities.Template
        | ExcelProviderCapabilities.Merge | ExcelProviderCapabilities.Async
        | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public void ExportEntity<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        ValidateEntityArguments(entity, layout, destination, cancellationToken);
        try
        {
            ExportEntityCore(entity, layout, destination, null, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML 实体导出失败。", exception,
                Provider, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <inheritdoc />
    public async Task ExportEntityAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        Stream destination, CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        ValidateEntityArguments(entity, layout, destination, cancellationToken);
        try
        {
            await using var buffer = new MemoryStream();
            await ExportEntityCoreAsync(entity, layout, buffer, null, cancellationToken)
                .ConfigureAwait(false);
            buffer.Position = 0;
            await buffer.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML 异步实体导出失败。", exception,
                Provider, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <inheritdoc />
    public void ExportEntityToFile<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout, string path,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        try
        {
            _fileExportCommitter.Commit(path,
                destination => ExportEntity(entity, layout, destination, cancellationToken),
                cancellationToken, Provider);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <inheritdoc />
    public async Task ExportEntityToFileAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        string path, CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        try
        {
            await _fileExportCommitter.CommitAsync(path,
                (destination, token) => ExportEntityAsync(entity, layout, destination, token),
                cancellationToken, Provider).ConfigureAwait(false);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <inheritdoc />
    public void ExportForTemplate<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        ExcelEntityTemplateOptions template, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var coreInvoked = false;
        try
        {
            ValidateEntityArguments(entity, layout, destination, cancellationToken);
            coreInvoked = true;
            ExportEntityCore(entity, layout, destination, template, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML 模板实体导出失败。", exception,
                Provider, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (!coreInvoked && !template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <inheritdoc />
    public async Task ExportForTemplateAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        ExcelEntityTemplateOptions template, Stream destination,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var coreInvoked = false;
        try
        {
            ValidateEntityArguments(entity, layout, destination, cancellationToken);
            coreInvoked = true;
            await using var buffer = new MemoryStream();
            await ExportEntityCoreAsync(entity, layout, buffer, template, cancellationToken)
                .ConfigureAwait(false);
            buffer.Position = 0;
            await buffer.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML 异步模板实体导出失败。", exception,
                Provider, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (!coreInvoked && !template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        try
        {
            ValidateArguments(request, destination, cancellationToken);
            ValidateRequest(request);
            using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Export);
            ExportWorkbookCore(request, destination, cancellationToken);
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
                "当前 ClosedXML 导出功能不受支持。", exception, Provider,
                BingOfficesOperation.Export, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML Excel 导出失败。", exception,
                Provider, BingOfficesStage.Write);
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
        try
        {
            ValidateArguments(request, destination, cancellationToken);
            ValidateRequest(request);
            await using var buffer = new MemoryStream();
            await using (var admission = await _admission
                .AcquireAsync(cancellationToken, BingOfficesOperation.Export)
                .ConfigureAwait(false))
            {
                ExportWorkbookCore(request, buffer, cancellationToken);
            }
            buffer.Position = 0;
            await buffer.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("ClosedXML 异步 Excel 导出失败。", exception,
                Provider, BingOfficesStage.Write);
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
                destination => Export(request, destination, cancellationToken), cancellationToken, Provider);
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
                (destination, token) => ExportAsync(request, destination, token), cancellationToken, Provider)
                .ConfigureAwait(false);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <summary>
    /// 在准入范围内同步执行实体工作簿导出。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="template">模板选项；为空时新建工作簿。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void ExportEntityCore<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        Stream destination, ExcelEntityTemplateOptions template, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        try
        {
            using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Export);
            ExportEntityWorkbookCore(entity, layout, destination, template, cancellationToken);
        }
        finally
        {
            if (template != null && !template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <summary>
    /// 在准入范围内执行实体工作簿导出，并异步复制外围流。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="template">模板选项；为空时新建工作簿。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private async Task ExportEntityCoreAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        Stream destination, ExcelEntityTemplateOptions template, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        try
        {
            await using var admission = await _admission
                .AcquireAsync(cancellationToken, BingOfficesOperation.Export)
                .ConfigureAwait(false);
            ExportEntityWorkbookCore(entity, layout, destination, template, cancellationToken);
        }
        finally
        {
            if (template != null && !template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <summary>
    /// 创建或读取工作簿并写入实体布局。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="template">模板选项；为空时新建工作簿。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void ExportEntityWorkbookCore<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        Stream destination, ExcelEntityTemplateOptions template, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        XLWorkbook workbook = null;
        try
        {
            if (template == null)
                workbook = new XLWorkbook();
            else
            {
                if (!template.Template.CanRead)
                    throw new ArgumentException("实体模板流不可读取。", nameof(template));
                ClosedXmlTemplatePreflight.Validate(template.Template);
                if (template.Template.CanSeek)
                    template.Template.Position = 0;
                workbook = new XLWorkbook(template.Template);
            }
            _entityExecutor.Write(workbook, entity, layout, template != null, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            workbook.SaveAs(destination);
            cancellationToken.ThrowIfCancellationRequested();
        }
        finally
        {
            workbook?.Dispose();
        }
    }

    /// <summary>
    /// 创建或读取工作簿并执行所有 Sheet 导出计划。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void ExportWorkbookCore(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken)
    {
        using var workbook = CreateWorkbook(request);
        ApplyMetadata(workbook, request);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheetRequest in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateSheetName(sheetRequest.Name);
            if (!names.Add(sheetRequest.Name))
                throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheetRequest.Name}");
            WriteSheet(workbook, sheetRequest, request.Template != null, cancellationToken);
        }
        cancellationToken.ThrowIfCancellationRequested();
        workbook.SaveAs(destination);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 校验实体导出的公共参数。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="entity">待导出的实体。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidateEntityArguments<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
        Stream destination, CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (entity == null)
            throw new ArgumentNullException(nameof(entity));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 按请求创建新工作簿或从模板加载工作簿。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <returns>待写入的 ClosedXML 工作簿。</returns>
    private XLWorkbook CreateWorkbook(ExcelWorkbookExportRequest request)
    {
        if (request.Template == null)
            return new XLWorkbook();
        if (!request.Template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(request));
        if (request.Template.CanSeek)
            request.Template.Position = 0;
        return new XLWorkbook(request.Template);
    }

    /// <summary>
    /// 将一个 Sheet 请求写入工作簿。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="request">Sheet 导出请求。</param>
    /// <param name="isTemplate">是否按模板模式写入。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void WriteSheet(XLWorkbook workbook, ExcelSheetExportRequest request, bool isTemplate,
        CancellationToken cancellationToken)
    {
        if (request.Data == null)
            throw new ArgumentNullException(nameof(request.Data));
        if (request.HeaderRowIndex < 0 || request.DataRowStartIndex <= request.HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(request.DataRowStartIndex));
        if (request.TemplateRegion != null)
            throw Unsupported("ClosedXML 第一版暂不支持通过命名区域定位模板区域。", BingOfficesStage.Preflight);

        var mappingConfiguration = ClosedXmlMappingPlanBuilder.MergeRequestDynamicColumns(
            request.MappingConfiguration, request.DynamicColumns);
        var plan = _planBuilder.Create(request.ItemType,
            request.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
            mappingConfiguration, MappingDirection.Export);
        var physicalConfiguration = MappingConfigurationMerger.Merge(
            request.MappingDocument?.Export, mappingConfiguration, MappingSourceKind.Request);
        var columns = ClosedXmlExportColumnPlanner.Create(request.ItemType, plan, request.DynamicColumns,
            physicalConfiguration);
        if (columns.Count == 0)
            throw new BingOfficesConfigurationException("ClosedXML 导出没有可写入的列。", stage: BingOfficesStage.Plan);

        var worksheet = FindWorksheet(workbook, request.Name);
        if (worksheet == null)
        {
            if (isTemplate)
                throw new BingOfficesConfigurationException($"模板缺少请求的 Sheet: {request.Name}",
                    stage: BingOfficesStage.Plan);
            worksheet = workbook.Worksheets.Add(request.Name);
        }
        if (request.Hidden)
            worksheet.Visibility = XLWorksheetVisibility.Hidden;

        var culture = request.Culture ?? CultureInfo.InvariantCulture;
        var headerRow = request.HeaderRowIndex + 1;
        var dataRow = request.DataRowStartIndex + 1;

        ApplyHeader(worksheet, columns, headerRow, request.HeaderStyle,
            ClosedXmlStyleKeyResolver.Resolve(plan.Style?.HeaderStyleKey, true), plan.Style?.NumberFormat,
            request.SheetStyle,
            isTemplate, request.TemplateCellOverwritePolicy);
        WriteCustomHeaders(worksheet, request.HeaderRows, isTemplate,
            request.CommentConflictPolicy, request.TemplateCellOverwritePolicy);
        var rowNumber = dataRow;
        foreach (var item in request.Data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item == null)
                throw new ArgumentException($"Sheet {request.Name} 包含 null 数据项。", nameof(request));
            var dynamicValues = request.DynamicGetter?.Invoke(item);
            if (request.FailOnUnknownDynamicValues && dynamicValues != null)
            {
                var known = new HashSet<string>(columns.Where(column => column.Dynamic != null)
                    .Select(column => column.Dynamic.Key),
                    StringComparer.OrdinalIgnoreCase);
                var unknown = dynamicValues.Keys.FirstOrDefault(key => !known.Contains(key));
                if (unknown != null)
                    throw new BingOfficesConfigurationException($"动态列值未定义: {unknown}",
                        stage: BingOfficesStage.Plan);
            }
            for (var index = 0; index < columns.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var binding = columns[index];
                var physicalColumn = binding.PhysicalColumnIndex + 1;
                var cell = worksheet.Cell(rowNumber, physicalColumn);
                PrepareTemplateCell(cell, isTemplate, request.TemplateCellOverwritePolicy);
                object value;
                if (binding.Fixed != null)
                    value = ClosedXmlValueAdapter.ConvertTo(binding.Property.GetValue(item), binding.Fixed,
                        binding.Property, request.Name, rowNumber, physicalColumn, culture);
                else
                {
                    object raw = null;
                    dynamicValues?.TryGetValue(binding.Dynamic.Key, out raw);
                    value = ClosedXmlValueAdapter.ConvertDynamicTo(raw, binding.Dynamic,
                        request.Name, rowNumber, physicalColumn, culture);
                }
                WriteValue(cell, value);
                ClosedXmlStyleAdapter.Apply(cell, request.SheetStyle);
                ClosedXmlStyleAdapter.Apply(cell,
                    ClosedXmlStyleKeyResolver.Resolve(plan.Style?.BodyStyleKey, false));
                ClosedXmlStyleAdapter.ApplyNumberFormat(cell, binding.MappingNumberFormat);
                ClosedXmlStyleAdapter.Apply(cell, request.BodyStyle);
                ClosedXmlStyleAdapter.Apply(cell, binding.BodyStyle);
                ClosedXmlStyleAdapter.ApplyNumberFormat(cell, binding.NumberFormat);
            }
            rowNumber++;
        }
        MergeDeclaredColumns(worksheet, columns, dataRow, rowNumber - 1);
        ApplyRowHeights(worksheet, request.RowHeight, headerRow, dataRow, rowNumber - 1,
            request.HeaderRows);
        ApplyColumnWidth(worksheet, columns, request.ColumnWidth);
    }

    /// <summary>
    /// 按映射属性上的合并声明合并连续数据单元格。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="columns">物理导出列。</param>
    /// <param name="firstRow">首个数据行号。</param>
    /// <param name="lastRow">最后数据行号。</param>
    private static void MergeDeclaredColumns(IXLWorksheet worksheet,
        IReadOnlyList<ClosedXmlExportColumn> columns, int firstRow, int lastRow)
    {
        if (lastRow <= firstRow)
            return;
        for (var index = 0; index < columns.Count; index++)
        {
            var property = columns[index].Property;
            if (property == null || property.GetCustomAttribute<MergeColumnsAttribute>() == null)
                continue;
            var groupStart = firstRow;
            var physicalColumn = columns[index].PhysicalColumnIndex + 1;
            var groupValue = worksheet.Cell(groupStart, physicalColumn).GetString();
            for (var row = firstRow + 1; row <= lastRow + 1; row++)
            {
                var current = row <= lastRow ? worksheet.Cell(row, physicalColumn).GetString() : null;
                if (row <= lastRow && string.Equals(groupValue, current, StringComparison.Ordinal))
                    continue;
                if (!string.IsNullOrWhiteSpace(groupValue) && row - groupStart > 1)
                    worksheet.Range(groupStart, physicalColumn, row - 1, physicalColumn).Merge();
                groupStart = row;
                groupValue = current;
            }
        }
    }

    /// <summary>
    /// 写入属性表头并按优先级应用表头样式。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="columns">物理导出列。</param>
    /// <param name="row">表头行号。</param>
    /// <param name="style">请求表头样式。</param>
    /// <param name="mappingStyle">映射表头样式。</param>
    /// <param name="mappingNumberFormat">映射数字格式。</param>
    /// <param name="sheetStyle">工作表样式。</param>
    /// <param name="isTemplate">是否按模板模式写入。</param>
    /// <param name="overwritePolicy">模板单元格覆盖策略。</param>
    private static void ApplyHeader(IXLWorksheet worksheet, IReadOnlyList<ClosedXmlExportColumn> columns,
        int row, Styles.ExcelCellStyle style, Styles.ExcelCellStyle mappingStyle,
        string mappingNumberFormat, Styles.ExcelCellStyle sheetStyle, bool isTemplate,
        ExcelTemplateCellOverwritePolicy overwritePolicy)
    {
        foreach (var column in columns)
        {
            var cell = worksheet.Cell(row, column.PhysicalColumnIndex + 1);
            PrepareTemplateCell(cell, isTemplate, overwritePolicy);
            cell.Value = column.Title;
            ClosedXmlStyleAdapter.Apply(cell, sheetStyle);
            ClosedXmlStyleAdapter.Apply(cell, mappingStyle);
            ClosedXmlStyleAdapter.ApplyNumberFormat(cell, mappingNumberFormat);
            ClosedXmlStyleAdapter.Apply(cell, style);
            ClosedXmlStyleAdapter.Apply(cell, column.HeaderStyle);
        }
    }

    /// <summary>
    /// 写入请求指定的自定义表头、批注和合并区域。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="headerRows">自定义表头行。</param>
    /// <param name="isTemplate">是否按模板模式写入。</param>
    /// <param name="commentPolicy">批注冲突策略。</param>
    /// <param name="overwritePolicy">模板单元格覆盖策略。</param>
    private static void WriteCustomHeaders(IXLWorksheet worksheet,
        IReadOnlyList<ExcelHeaderRow> headerRows, bool isTemplate,
        ExcelCommentConflictPolicy commentPolicy, ExcelTemplateCellOverwritePolicy overwritePolicy)
    {
        foreach (var headerRow in headerRows ?? Array.Empty<ExcelHeaderRow>())
        {
            foreach (var headerCell in headerRow.Cells)
            {
                var row = headerRow.RowIndex + 1;
                var column = headerCell.ColumnIndex + 1;
                var cell = worksheet.Cell(row, column);
                PrepareTemplateCell(cell, isTemplate, overwritePolicy);
                WriteValue(cell, headerCell.Value);
                if (headerCell.Comment != null)
                    ApplyComment(cell, headerCell.Comment, commentPolicy);
                if (headerCell.RowSpan > 1 || headerCell.ColumnSpan > 1)
                    worksheet.Range(row, column, row + headerCell.RowSpan - 1,
                        column + headerCell.ColumnSpan - 1).Merge();
            }
        }
    }

    /// <summary>
    /// 按模板覆盖策略准备目标单元格。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="isTemplate">是否按模板模式写入。</param>
    /// <param name="overwritePolicy">模板单元格覆盖策略。</param>
    private static void PrepareTemplateCell(IXLCell cell, bool isTemplate,
        ExcelTemplateCellOverwritePolicy overwritePolicy)
    {
        if (!isTemplate || overwritePolicy != ExcelTemplateCellOverwritePolicy.ReplaceTemplate)
            return;
        cell.Clear(XLClearOptions.Contents | XLClearOptions.AllFormats
            | XLClearOptions.Comments);
    }

    /// <summary>
    /// 按批注冲突策略写入单元格批注。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="comment">待写入批注。</param>
    /// <param name="conflictPolicy">批注冲突策略。</param>
    private static void ApplyComment(IXLCell cell, ExcelComment comment,
        ExcelCommentConflictPolicy conflictPolicy)
    {
        if (cell.HasComment)
        {
            switch (conflictPolicy)
            {
                case ExcelCommentConflictPolicy.Preserve:
                    return;
                case ExcelCommentConflictPolicy.Fail:
                    throw new BingOfficesConfigurationException($"单元格已有批注: {cell.Address}",
                        stage: BingOfficesStage.Plan);
                case ExcelCommentConflictPolicy.Append:
                    var existing = cell.GetComment().ToString();
                    comment = new ExcelComment(existing + Environment.NewLine + comment.Text,
                        string.IsNullOrWhiteSpace(comment.Author) ? cell.GetComment().Author : comment.Author,
                        comment.Visible);
                    cell.GetComment().Delete();
                    break;
                case ExcelCommentConflictPolicy.Replace:
                    cell.GetComment().Delete();
                    break;
            }
        }
        var created = cell.CreateComment();
        created.AddText(comment.Text);
        created.SetAuthor(comment.Author);
        created.SetVisible(comment.Visible);
    }

    /// <summary>
    /// 按列宽配置调整实际物理列。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="columns">物理导出列。</param>
    /// <param name="options">列宽配置。</param>
    private static void ApplyColumnWidth(IXLWorksheet worksheet,
        IReadOnlyList<ClosedXmlExportColumn> columns, ExcelColumnWidthOptions options)
    {
        if (options == null || options.Mode == ExcelColumnWidthMode.None)
            return;
        options.Validate();
        foreach (var physicalIndex in columns.Select(column => column.PhysicalColumnIndex).Distinct().OrderBy(index => index))
        {
            var column = worksheet.Column(physicalIndex + 1);
            if (options.Mode == ExcelColumnWidthMode.Fixed)
                column.Width = options.FixedWidth.Value;
            else
                column.AdjustToContents();
            if (options.MinWidth.HasValue && column.Width < options.MinWidth.Value)
                column.Width = options.MinWidth.Value;
            if (options.MaxWidth.HasValue && column.Width > options.MaxWidth.Value)
                column.Width = options.MaxWidth.Value;
        }
    }

    /// <summary>
    /// 按行高配置设置表头和数据行高度。
    /// </summary>
    /// <param name="worksheet">目标工作表。</param>
    /// <param name="options">行高配置。</param>
    /// <param name="headerRow">属性表头行号。</param>
    /// <param name="dataRow">首个数据行号。</param>
    /// <param name="lastDataRow">最后数据行号。</param>
    /// <param name="customHeaders">自定义表头行。</param>
    private static void ApplyRowHeights(IXLWorksheet worksheet, ExcelRowHeightOptions options,
        int headerRow, int dataRow, int lastDataRow, IReadOnlyList<ExcelHeaderRow> customHeaders)
    {
        if (options == null)
            return;
        options.Validate();
        if (options.HeaderHeight.HasValue)
        {
            worksheet.Row(headerRow).Height = options.HeaderHeight.Value;
            foreach (var row in customHeaders ?? Array.Empty<ExcelHeaderRow>())
                worksheet.Row(row.RowIndex + 1).Height = options.HeaderHeight.Value;
        }
        if (options.BodyHeight.HasValue && lastDataRow >= dataRow)
        {
            for (var row = dataRow; row <= lastDataRow; row++)
                worksheet.Row(row).Height = options.BodyHeight.Value;
        }
    }

    /// <summary>
    /// 将公共值或公式文本写入 ClosedXML 单元格。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="value">待写入的值。</param>
    private static void WriteValue(IXLCell cell, object value)
    {
        if (value is string formulaText && formulaText.StartsWith("=", StringComparison.Ordinal))
        {
            cell.FormulaA1 = formulaText.Substring(1);
            return;
        }
        switch (value)
        {
            case null: cell.Clear(XLClearOptions.Contents); break;
            case string stringValue: cell.Value = stringValue; break;
            case bool boolean: cell.Value = boolean; break;
            case DateTime date: cell.Value = date; break;
            case DateTimeOffset offset: cell.Value = offset.DateTime; break;
            case byte number: cell.Value = number; break;
            case sbyte number: cell.Value = number; break;
            case short number: cell.Value = number; break;
            case ushort number: cell.Value = number; break;
            case int number: cell.Value = number; break;
            case uint number: cell.Value = number; break;
            case long number: cell.Value = number; break;
            case ulong number: cell.Value = number; break;
            case float number: cell.Value = number; break;
            case double number: cell.Value = number; break;
            case decimal number: cell.Value = number; break;
            default: cell.Value = Convert.ToString(value, CultureInfo.InvariantCulture); break;
        }
    }

    /// <summary>
    /// 按名称查找工作表。
    /// </summary>
    /// <param name="workbook">工作簿。</param>
    /// <param name="name">工作表名称。</param>
    /// <returns>匹配的工作表；未找到时返回 <see langword="null" />。</returns>
    private static IXLWorksheet FindWorksheet(XLWorkbook workbook, string name) =>
        workbook.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, name,
            StringComparison.OrdinalIgnoreCase));

    /// <summary>
    /// 校验 Workbook 导出请求的 ClosedXML 支持边界。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    private static void ValidateRequest(ExcelWorkbookExportRequest request)
    {
        if (request.Format != ExcelFormat.Xlsx)
            throw Unsupported("ClosedXML Provider 仅支持 XLSX。", BingOfficesStage.Preflight);
        if (request.Template != null)
            ClosedXmlTemplatePreflight.Validate(request.Template);
        if (request.Charts().Count > 0)
            throw Unsupported("ClosedXML Provider 第一版不创建图表。", BingOfficesStage.Preflight);
        foreach (var sheet in request.Sheets)
        {
            if (sheet.Charts?.Count > 0)
                throw Unsupported("ClosedXML Provider 第一版不创建图表。", BingOfficesStage.Preflight);
        }
    }

    /// <summary>
    /// 校验导出请求、目标流和取消状态。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
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
    /// 校验工作表名称符合 XLSX 限制。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    private static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31)
            throw new ArgumentException("工作表名称不能超过 31 个字符。", nameof(name));
        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称包含非法字符: {name}", nameof(name));
    }

    /// <summary>
    /// 将请求中的元数据写入工作簿属性。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="request">Workbook 导出请求。</param>
    private static void ApplyMetadata(XLWorkbook workbook, ExcelWorkbookExportRequest request)
    {
        if (!request.MetadataSpecified || request.Metadata == null)
            return;
        workbook.Properties.Author = request.Metadata.Author;
        workbook.Properties.Company = request.Metadata.Company;
        workbook.Properties.Title = request.Metadata.Title;
        workbook.Properties.Subject = request.Metadata.Subject;
        workbook.Properties.Category = request.Metadata.Category;
        workbook.Properties.Comments = request.Metadata.Description;
    }

    /// <summary>
    /// 创建带 ClosedXML 导出上下文的 Unsupported 异常。
    /// </summary>
    /// <param name="message">异常消息。</param>
    /// <param name="stage">异常阶段。</param>
    /// <returns>结构化 Unsupported 异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message, BingOfficesStage stage) =>
        new(message, provider: Provider, operation: BingOfficesOperation.Export, stage: stage);

    /// <summary>
    /// 按请求所有权设置释放模板流。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    private static void DisposeTemplate(ExcelWorkbookExportRequest request)
    {
        if (request?.Template != null && !request.LeaveTemplateOpen)
            request.Template.Dispose();
    }
}

/// <summary>
/// 将公共样式键解析为 ClosedXML Provider 支持的样式。
/// </summary>
internal static class ClosedXmlStyleKeyResolver
{
    /// <summary>
    /// 解析表头或正文样式键。
    /// </summary>
    /// <param name="key">样式键。</param>
    /// <param name="header">是否解析表头样式。</param>
    /// <returns>解析后的公共样式；键为空时返回 <see langword="null" />。</returns>
    internal static Styles.ExcelCellStyle Resolve(string key, bool header)
    {
        if (string.IsNullOrWhiteSpace(key))
            return null;
        if (string.Equals(key, "header", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "bold", StringComparison.OrdinalIgnoreCase))
            return new Styles.ExcelCellStyle { Bold = true };
        if (string.Equals(key, "body", StringComparison.OrdinalIgnoreCase)
            || string.Equals(key, "default", StringComparison.OrdinalIgnoreCase))
            return new Styles.ExcelCellStyle();
        throw new BingOfficesConfigurationException($"未注册的{(header ? "表头" : "正文")}样式键: {key}",
            stage: BingOfficesStage.Plan);
    }
}

/// <summary>
/// 提供 ClosedXML 导出请求的 Provider 辅助访问器。
/// </summary>
internal static class ClosedXmlExportRequestExtensions
{
    /// <summary>
    /// 汇总 Workbook 及其各 Sheet 的图表定义。
    /// </summary>
    /// <param name="request">Workbook 导出请求。</param>
    /// <returns>图表定义集合。</returns>
    public static IReadOnlyList<ExcelChartDefinition> Charts(this ExcelWorkbookExportRequest request) =>
        request.Sheets.SelectMany(sheet => sheet.Charts ?? Array.Empty<ExcelChartDefinition>()).ToArray();
}
