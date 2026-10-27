using System.Reflection;
using System.Runtime.ExceptionServices;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using Bing.Offices.Streaming;
using Bing.Offices.Styles;
using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;

namespace Bing.Offices.Exports;

/// <summary>显式选择的 SXSSF 导出实现；每次调用独立拥有工作簿和临时会话。</summary>
internal sealed class NpoiExcelStreamingExporter : IExcelStreamingExporter
{
    private readonly NpoiExportPlanBuilder _plans;
    private readonly IFileExportCommitter _committer;
    private readonly BingOfficesExceptionDispatcher _exceptions;

    internal NpoiExcelStreamingExporter(IEnumerable<IExcelValueConverter> converters = null,
        IExcelMappingPlanFactory mapping = null, IEnumerable<IBingOfficesExceptionObserver> observers = null,
        IFileExportCommitter committer = null)
    {
        _plans = new NpoiExportPlanBuilder(mapping ?? NpoiMappingPlanFactoryResolver.CreateDefault(converters));
        _committer = committer ?? new DefaultFileExportCommitter();
        _exceptions = new BingOfficesExceptionDispatcher(observers);
    }

    public void Export(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
        => ExportCore(request, destination, options, cancellationToken, false);

    /// <summary>文件出口需为提交副本预留额度；调用方流不计入库临时预算。</summary>
    private void ExportCore(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options, CancellationToken cancellationToken, bool reserveCommitCopy)
    {
        options ??= new();
        ValidateDestination(destination);
        var writes = Prepare(request, options, cancellationToken);
        var session = new ExcelStreamingTempSession(options.TempDirectory, options.MaxTemporaryBytes);
        Exception failure = null;
        try
        {
            var staged = session.CreateFile();
            Generate(request, options, writes, staged, cancellationToken);
            if (reserveCommitCopy) session.ReserveCommitCopy(staged.Length);
            staged.Position = 0;
            var buffer = new byte[65536];
            int read;
            while ((read = staged.Read(buffer, 0, buffer.Length)) != 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                destination.Write(buffer, 0, read);
            }
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (Exception e) { failure = e; Observe(e); throw; }
        finally { StreamingCleanup.Run(session.Dispose, failure); }
    }

    public Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
        => ExportCoreAsync(request, destination, options, cancellationToken, false);

    /// <summary>真实异步复制前，将文件提交目标与暂存包同时计入本次额度。</summary>
    private async Task ExportCoreAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options, CancellationToken cancellationToken, bool reserveCommitCopy)
    {
        options ??= new();
        ValidateDestination(destination);
        var writes = Prepare(request, options, cancellationToken);
        var session = new ExcelStreamingTempSession(options.TempDirectory, options.MaxTemporaryBytes);
        Exception failure = null;
        try
        {
            var staged = session.CreateFile();
            Generate(request, options, writes, staged, cancellationToken);
            await staged.FlushAsync(cancellationToken).ConfigureAwait(false);
            if (reserveCommitCopy) session.ReserveCommitCopy(staged.Length);
            staged.Position = 0;
            await staged.CopyToAsync(destination, 65536, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (Exception e) { failure = e; Observe(e); throw; }
        finally { StreamingCleanup.Run(session.Dispose, failure); }
    }

    public void ExportToFile(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
    {
        try { _committer.Commit(path, stream => ExportCore(request, stream, options, cancellationToken, true), cancellationToken, "XLSX"); }
        catch (BingOfficesException e) { _exceptions.Observe(e); throw; }
    }

    public async Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
    {
        try
        {
            await _committer.CommitAsync(path, (stream, token) => ExportCoreAsync(request, stream, options, token, true),
                cancellationToken, "XLSX").ConfigureAwait(false);
        }
        catch (BingOfficesException e) { _exceptions.Observe(e); throw; }
    }

    private void Observe(Exception e) { if (e is BingOfficesException office) _exceptions.Observe(office); }
    private static void ValidateDestination(Stream stream)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanWrite) throw new ArgumentException("目标流不可写。", nameof(stream));
    }

    private IReadOnlyList<Action<SXSSFWorkbook>> Prepare(ExcelWorkbookExportRequest request,
        ExcelStreamingExportOptions options, CancellationToken token)
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        options.Validate();
        token.ThrowIfCancellationRequested();
        if (request.Template != null || request.Format != ExcelFormat.Xlsx)
            throw new NotSupportedException("低内存导出只支持新建 XLSX，不支持模板。");
        if (request.Sheets.Count > options.MaxSheets) throw new StreamingResourceLimitException("Sheet 数量超过限制。");
        var metadata = request.Metadata;
        if (new[] { metadata.Author, metadata.Company, metadata.Title, metadata.Subject, metadata.Category, metadata.Description }
            .Sum(value => (long)(value?.Length ?? 0)) > options.MaxRowCharacters)
            throw new StreamingResourceLimitException("工作簿元数据超过文本限制。");
        var maps = _plans.Create(request);
        var writes = new List<Action<SXSSFWorkbook>>();
        foreach (var sheet in request.Sheets)
        {
            if (sheet.Charts.Count != 0 || sheet.HeaderRows.Count != 0 || sheet.TemplateRegion != null
                || sheet.ColumnWidth?.Mode is ExcelColumnWidthMode.AutoFit or ExcelColumnWidthMode.Adaptive)
                throw new NotSupportedException("低内存导出不支持图表、多级表头、模板区域或自动列宽。");
            if (sheet.DataRowStartIndex > 1048575 || sheet.HeaderRowIndex > 1048575)
                throw new ArgumentOutOfRangeException(nameof(request), "行位置超过 XLSX 格式范围。");
            try
            {
                writes.Add((Action<SXSSFWorkbook>)typeof(NpoiExcelStreamingExporter)
                    .GetMethod(nameof(PrepareSheet), BindingFlags.Static | BindingFlags.NonPublic)!
                    .MakeGenericMethod(sheet.ItemType).Invoke(null, new object[] { sheet, maps[sheet], options, token }));
            }
            catch (TargetInvocationException e) when (e.InnerException != null)
            { ExceptionDispatchInfo.Capture(e.InnerException).Throw(); throw; }
        }
        return writes;
    }

    private static Action<SXSSFWorkbook> PrepareSheet<T>(ExcelSheetExportRequest request, IExcelMappingPlan map,
        ExcelStreamingExportOptions options, CancellationToken token) where T : class, new()
    {
        var definitions = map.DynamicColumns.Select(column => NpoiExportColumnPlanner.CreateDynamicDefinition(column,
            request.DynamicColumns.FirstOrDefault(item => string.Equals(item.Key, column.Key, StringComparison.OrdinalIgnoreCase)), map)).ToArray();
        NpoiExportColumnPlanner.ValidateDynamicDefinitions(definitions);
        var columns = NpoiExportColumnPlanner.CreateColumns<T>(map, definitions);
        NpoiExportColumnPlanner.ValidateDynamicColumns(definitions, columns);
        if (columns.Count > options.MaxColumns) throw new StreamingResourceLimitException("列数超过限制。");
        if (columns.Any(column => column.IsMerged || NpoiImportRowMaterializer.IsImageColumn(column)))
            throw new NotSupportedException("低内存导出不支持合并列或图片列。");
        var headerStyle = NpoiExportSheetWriter.ResolveStyle(map.Style?.HeaderStyleKey, true);
        var bodyStyle = NpoiExportSheetWriter.ResolveStyle(map.Style?.BodyStyleKey, false);
        foreach (var column in columns)
            if (column.Title.Length > options.MaxCellCharacters) throw new StreamingResourceLimitException("表头文本超过限制。");
        if (columns.Sum(column => (long)column.Title.Length) > options.MaxRowCharacters)
            throw new StreamingResourceLimitException("表头总文本超过限制。");
        return book => WriteSheet<T>(book, request, columns, headerStyle, bodyStyle, options, token);
    }

    private static void WriteSheet<T>(SXSSFWorkbook book, ExcelSheetExportRequest request,
        IReadOnlyList<ExcelColumnPlan> columns, ExcelCellStyle mappingHeader, ExcelCellStyle mappingBody,
        ExcelStreamingExportOptions options, CancellationToken token) where T : class, new()
    {
        var sheet = (SXSSFSheet)book.CreateSheet(request.Name);
        if (request.Hidden) book.SetSheetVisibility(book.GetSheetIndex(sheet), SheetVisibility.Hidden);
        NpoiExportSheetWriter.ApplyColumnWidths(sheet, columns, 0, 0, request.ColumnWidth);
        var headerAttribute = typeof(T).GetCustomAttributes(typeof(HeaderAttribute), false).Cast<HeaderAttribute>().SingleOrDefault();
        var wrap = typeof(T).IsDefined(typeof(WrapTextAttribute), false);
        var header = sheet.CreateRow(request.HeaderRowIndex);
        for (var i = 0; i < columns.Count; i++)
        {
            var cell = header.CreateCell(i);
            cell.SetCellValue(columns[i].Title);
            cell.CellStyle = NpoiStyleCache.Compose(book, cell.CellStyle,
                columns[i].HeaderStyle ?? request.HeaderStyle ?? mappingHeader ?? request.SheetStyle);
            if (headerAttribute != null) cell.CellStyle = NpoiStyleCache.ApplyHeaderAttribute(book, cell.CellStyle, headerAttribute);
            if (wrap) cell.CellStyle = NpoiStyleCache.Compose(book, cell.CellStyle, new ExcelCellStyle { WrapText = true });
        }
        CheckStyles(book, options);
        // 表头不需要回访，立刻释放。正文最多保留窗口行数和窗口字符额度。
        sheet.FlushRows();
        var window = new Queue<int>();
        long characters = 0;
        var rowIndex = request.DataRowStartIndex;
        var keys = new HashSet<string>(columns.Where(c => c.IsDynamic).Select(c => c.Key), StringComparer.Ordinal);
        foreach (T item in request.Data)
        {
            token.ThrowIfCancellationRequested();
            if (rowIndex > 1048575) throw new StreamingResourceLimitException("数据行超过 XLSX 格式范围。");
            // 为当前行预留完整行额度，避免先分配一行后才发现窗口文本超额。
            if (characters > options.MaxBufferedCharacters - options.MaxRowCharacters)
            {
                sheet.FlushRows();
                window.Clear();
                characters = 0;
            }
            var row = sheet.CreateRow(rowIndex);
            var values = NpoiExportSheetWriter.ReadDynamicValues(request, item, columns, request.Name, rowIndex + 1);
            NpoiExportSheetWriter.ValidateUnknownDynamicValues(request, values, keys);
            var rowCharacters = 0;
            for (var i = 0; i < columns.Count; i++)
            {
                token.ThrowIfCancellationRequested();
                var cell = row.CreateCell(i);
                NpoiExportSheetWriter.WriteCell(cell, item, columns[i], values, request.Name, rowIndex + 1, i + 1, request.Culture);
                var length = cell.CellType == CellType.String ? cell.StringCellValue.Length : 0;
                if (length > options.MaxCellCharacters || length > options.MaxRowCharacters - rowCharacters)
                    throw new StreamingResourceLimitException("单元格或单行文本超过限制。");
                rowCharacters += length;
                cell.CellStyle = NpoiStyleCache.Compose(book, cell.CellStyle,
                    columns[i].BodyStyle ?? request.BodyStyle ?? mappingBody ?? request.SheetStyle);
                if (wrap) cell.CellStyle = NpoiStyleCache.Compose(book, cell.CellStyle, new ExcelCellStyle { WrapText = true });
                CheckStyles(book, options);
            }
            window.Enqueue(rowCharacters); characters += rowCharacters;
            while (window.Count > options.RowWindowSize) characters -= window.Dequeue();
            rowIndex++;
        }
        sheet.FlushRows();
    }

    private static void CheckStyles(SXSSFWorkbook book, ExcelStreamingExportOptions options)
    {
        if (book.NumCellStyles > options.MaxStyles || book.NumberOfFonts > options.MaxFonts)
            throw new StreamingResourceLimitException("工作簿样式或字体数量超过限制。");
    }

    private static void Generate(ExcelWorkbookExportRequest request, ExcelStreamingExportOptions options,
        IReadOnlyList<Action<SXSSFWorkbook>> writes, Stream staged, CancellationToken token)
    {
        var book = new SXSSFWorkbook(options.RowWindowSize) { CompressTempFiles = options.CompressTemporaryFiles };
        Exception failure = null;
        try
        {
            ExcelHelper.ApplyWorkbookMetadata(book.XssfWorkbook, request.Metadata);
            foreach (var write in writes) { token.ThrowIfCancellationRequested(); write(book); }
            token.ThrowIfCancellationRequested();
            using var guarded = new StreamingWriteCancellationStream(staged, token);
            book.Write(guarded, true);
            token.ThrowIfCancellationRequested();
        }
        catch (Exception e) { failure = e; throw; }
        finally { StreamingCleanup.Run(() => DisposeWorkbook(book), failure); }
    }

    /// <summary>必须先刷出并删除 Sheet 文件，再关闭基础包。异常时继续尝试所有 writer。</summary>
    internal static void DisposeWorkbook(SXSSFWorkbook book)
    {
        var errors = new List<Exception>();
        try { if (!book.Dispose()) errors.Add(new IOException("NPOI 工作表临时文件删除失败。")); }
        catch (Exception e) { errors.Add(e); }
        if (errors.Count != 0)
            for (var i = 0; i < book.NumberOfSheets; i++)
                try
                {
                    if (!((SXSSFSheet)book.GetSheetAt(i)).SheetDataWriter.Dispose())
                        errors.Add(new IOException("NPOI writer 临时文件删除失败。"));
                }
                catch (Exception e) { errors.Add(e); }
        try { book.XssfWorkbook.Close(); } catch (Exception e) { errors.Add(e); }
        if (errors.Count != 0) throw new AggregateException("NPOI 工作簿清理失败。", errors);
    }
}
