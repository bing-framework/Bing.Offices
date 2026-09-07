using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Internals;
using Bing.Offices.Npoi.Resolvers;
using Bing.Offices.IO;
using System.Globalization;
using System.Collections.Concurrent;
using System.Reflection;
using System.Text.RegularExpressions;
using NPOI.SS.Util;
using NPOI.SS.UserModel.Charts;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 基于 NPOI 的单工作簿 Excel 导出器；NPOI 工作簿在内存中构建后写入目标流。
/// </summary>
internal sealed class NpoiExcelExporter : IExcelExporter
{
    private delegate void WriteSheetInvoker(NpoiExcelExporter target, NPOI.SS.UserModel.IWorkbook workbook,
        ExcelSheetExportRequest request, bool isTemplate, CancellationToken cancellationToken,
        IExcelMappingPlan mapping);

    private static readonly ConcurrentDictionary<Type, WriteSheetInvoker> WriteSheetInvokers = new();
    /// <summary>
    /// 当前导出器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>
    /// 将导出请求编译为按工作表执行的映射计划生成器。
    /// </summary>
    private readonly NpoiExportPlanBuilder _planBuilder;
    /// <summary>
    /// 根据列计划将实体数据写入 NPOI 工作表的写入器。
    /// </summary>
    private readonly NpoiExportSheetWriter _sheetWriter;
    /// <summary>负责原子文件提交的可替换 SPI。</summary>
    private readonly IFileExportCommitter _fileExportCommitter;
    /// <summary>观察并记录公共 Excel 导出异常。</summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个<see cref="NpoiExcelExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    /// <param name="fileExportCommitter">可替换的原子文件提交器。</param>
    public NpoiExcelExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null,
        IFileExportCommitter fileExportCommitter = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _planBuilder = new NpoiExportPlanBuilder(mappingPlanFactory ?? NpoiMappingPlanFactoryResolver.CreateDefault(
            _valueConverters));
        _sheetWriter = new NpoiExportSheetWriter();
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
    }

    /// <inheritdoc />
    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        CancellationToken cancellationToken = default)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();

        try
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheetRequest in request.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateSheetName(sheetRequest.Name);
                if (!names.Add(sheetRequest.Name))
                    throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheetRequest.Name}");
            }
            using var workbook = CreateWorkbook(request);
            Dictionary<ExcelSheetExportRequest, IExcelMappingPlan> planBySheet;
            try
            {
                planBySheet = _planBuilder.Create(request);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesConfigurationException("Excel 导出映射配置无效。", exception,
                    BingOfficesStage.Plan);
            }
            foreach (var sheetRequest in request.Sheets)
            {
                cancellationToken.ThrowIfCancellationRequested();
                WriteSheet(workbook, sheetRequest, request.Template != null, cancellationToken,
                    planBySheet[sheetRequest]);
            }

            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                workbook.Write(new NpoiNonDisposingStream(destination, cancellationToken), false);
            }
            catch (Exception exception) when (cancellationToken.IsCancellationRequested
                && exception.GetBaseException() is OperationCanceledException)
            {
                throw new OperationCanceledException(cancellationToken);
            }
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
            var translated = new BingOfficesUnsupportedFeatureException("当前 Excel 导出功能不受 NPOI 支持。",
                exception, "NPOI", BingOfficesOperation.Export, BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("Excel 导出失败。", exception, "NPOI",
                BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (request.Template != null && !request.LeaveTemplateOpen)
                request.Template.Dispose();
        }
    }

    /// <inheritdoc />
    public void ExportToFile(ExcelWorkbookExportRequest request, string path,
        CancellationToken cancellationToken = default)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));

        try
        {
            _fileExportCommitter.Commit(path,
                destination => Export(request, destination, cancellationToken),
                cancellationToken, "Excel");
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <summary>
    /// 创建普通或模板工作簿。模板加载后沿用同一 Sheet Writer。
    /// </summary>
    /// <param name="request">包含格式、模板流和元数据的导出请求。</param>
    /// <returns>已准备好的 NPOI 工作簿。</returns>
    private static NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExcelWorkbookExportRequest request)
    {
        if (request.Template == null)
            return ExcelHelper.PrepareWorkbook(request.Format, request.Metadata);
        if (!request.Template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(request));
        var workbook = NPOI.SS.UserModel.WorkbookFactory.Create(new NpoiNonDisposingStream(request.Template));
        if (request.MetadataSpecified)
            ExcelHelper.ApplyWorkbookMetadata(workbook, request.Metadata);
        return workbook;
    }

    /// <summary>
    /// 验证 Excel Sheet 名称边界。
    /// </summary>
    /// <param name="name">待验证的工作表名称。</param>
    private static void ValidateSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("工作表名称不能为空。", nameof(name));
        if (name.Length > 31)
            throw new ArgumentException("工作表名称不能超过 31 个字符。", nameof(name));
        if (name.IndexOfAny(new[] { ':', '\\', '/', '?', '*', '[', ']' }) >= 0)
            throw new ArgumentException($"工作表名称包含非法字符: {name}", nameof(name));
    }

    /// <summary>通过反射分派到对应实体类型的工作表写入方法。</summary>
    /// <param name="workbook">目标 NPOI 工作簿。</param>
    /// <param name="request">当前工作表导出请求。</param>
    /// <param name="isTemplate">是否基于模板工作簿导出。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="mapping">当前工作表的不可变映射计划。</param>
    private void WriteSheet(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken, IExcelMappingPlan mapping)
    {
        WriteSheetInvokers.GetOrAdd(request.ItemType, CreateWriteSheetInvoker)(this, workbook, request,
            isTemplate, cancellationToken, mapping);
    }

    private static WriteSheetInvoker CreateWriteSheetInvoker(Type itemType)
    {
        var method = typeof(NpoiExcelExporter).GetMethod(nameof(WriteTypedSheet),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(itemType);
        return (WriteSheetInvoker)method.CreateDelegate(typeof(WriteSheetInvoker));
    }

    /// <summary>
    /// 执行一个泛型 Sheet 的统一列计划和 Cell Writer。
    /// </summary>
    /// <typeparam name="T">工作表数据项类型。</typeparam>
    /// <param name="workbook">目标 NPOI 工作簿。</param>
    /// <param name="request">当前工作表导出请求。</param>
    /// <param name="isTemplate">是否基于模板工作簿导出。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="map">当前工作表的映射计划。</param>
    private void WriteTypedSheet<T>(NPOI.SS.UserModel.IWorkbook workbook, ExcelSheetExportRequest request,
        bool isTemplate, CancellationToken cancellationToken, IExcelMappingPlan map) where T : class, new()
    {
        var sheet = workbook.GetSheet(request.Name);
        if (sheet == null && isTemplate)
            throw new BingOfficesConfigurationException($"模板缺少请求的 Sheet: {request.Name}",
                stage: BingOfficesStage.Plan);
        sheet ??= workbook.CreateSheet(request.Name);
        if (request.Hidden)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), NPOI.SS.UserModel.SheetVisibility.Hidden);
        if (request.HeaderRowIndex < 0 || request.DataRowStartIndex <= request.HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(request.DataRowStartIndex));

        (int Row, int Column) templateOrigin;
        IReadOnlyList<ExcelColumnPlan> columns;
        try
        {
            templateOrigin = ResolveTemplateOrigin(workbook, sheet, request, isTemplate);
            var dynamicDefinitions = map.DynamicColumns.Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column,
                request.DynamicColumns.FirstOrDefault(item => string.Equals(item.Key, column.Key,
                    StringComparison.OrdinalIgnoreCase)), map)).ToArray();
            NpoiExportColumnPlanner.ValidateDynamicDefinitions(dynamicDefinitions);
            columns = NpoiExportColumnPlanner.CreateColumns<T>(map, dynamicDefinitions);
            NpoiExportColumnPlanner.ValidateDynamicColumns(dynamicDefinitions, columns);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException("Excel 导出映射配置无效。", exception,
                BingOfficesStage.Plan);
        }
        var headerRowIndex = templateOrigin.Row + request.HeaderRowIndex;
        var firstColumnIndex = templateOrigin.Column;
        _sheetWriter.Write<T>(workbook, request, cancellationToken, map, columns, templateOrigin.Row,
            firstColumnIndex);
    }

    /// <summary>
    /// 解析模板命名区域的起点；普通导出从零坐标开始。
    /// </summary>
    private static (int Row, int Column) ResolveTemplateOrigin(NPOI.SS.UserModel.IWorkbook workbook,
        NPOI.SS.UserModel.ISheet sheet, ExcelSheetExportRequest request, bool isTemplate)
    {
        if (!isTemplate || string.IsNullOrWhiteSpace(request.TemplateRegion))
            return (0, 0);
        var name = workbook.GetName(request.TemplateRegion);
        if (name == null || string.IsNullOrWhiteSpace(name.RefersToFormula))
            throw new InvalidOperationException($"模板缺少命名区域: {request.TemplateRegion}");
        var formula = name.RefersToFormula.TrimStart('=');
        var separator = formula.LastIndexOf('!');
        if (separator < 0)
            throw new InvalidOperationException($"模板命名区域缺少 Sheet 引用: {request.TemplateRegion}");
        var sheetName = formula.Substring(0, separator).Trim('\'', ' ');
        if (!string.Equals(sheetName, sheet.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException($"模板命名区域不属于请求 Sheet: {request.TemplateRegion}");
        var address = formula.Substring(separator + 1).Split(':')[0].Replace("$", string.Empty);
        var match = Regex.Match(address, "^([A-Za-z]+)([0-9]+)$");
        if (!match.Success)
            throw new InvalidOperationException($"模板命名区域地址无效: {request.TemplateRegion}");
        var column = 0;
        foreach (var character in match.Groups[1].Value.ToUpperInvariant())
            column = column * 26 + character - 'A' + 1;
        return (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1, column - 1);
    }


}
