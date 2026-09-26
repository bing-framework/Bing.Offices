using System.Diagnostics;
using System.Globalization;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// ClosedXML 导入失败工作簿生成器。
/// </summary>
/// <remarks>先完整序列化到临时产物，再由产物提交到调用方流。</remarks>
internal static class ClosedXmlFailureWorkbookWriter
{
    /// <summary>
    /// ClosedXML Provider 名称。
    /// </summary>
    private const string Provider = "ClosedXML";

    /// <summary>
    /// 根据导入错误创建已完整序列化的失败工作簿产物。
    /// </summary>
    /// <param name="source">原始导入工作簿。</param>
    /// <param name="options">失败工作簿选项。</param>
    /// <param name="errors">结构化导入错误。</param>
    /// <param name="resolvedSheetRequests">实际工作表名称对应的导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>可提交的临时产物；不需要输出时返回 null。</returns>
    internal static ClosedXmlFailureWorkbookArtifact Create(XLWorkbook source,
        ExcelImportFailureOptions options, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return null;
        cancellationToken.ThrowIfCancellationRequested();
        ValidatePreflight(source, options, errors, resolvedSheetRequests);

        XLWorkbook output = source;
        XLWorkbook independent = null;
        try
        {
            if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly)
                output = independent = CreateErrorRowsWorkbook(source, errors, resolvedSheetRequests,
                    cancellationToken);
            else if (options.Mode == ExcelImportFailureWorkbookMode.AnnotatedOriginal)
                AnnotateErrors(output, errors, options.CommentConflictPolicy, cancellationToken);
            else
                throw new ArgumentOutOfRangeException(nameof(options.Mode));
            WriteSummary(output, errors, cancellationToken);
            return Serialize(output, options, cancellationToken);
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
            throw new BingOfficesImportException("ClosedXML 失败工作簿内容生成失败。", exception,
                Provider, BingOfficesStage.Write);
        }
        finally
        {
            independent?.Dispose();
        }
    }

    /// <summary>
    /// 在修改源工作簿前校验失败输出资源预算。
    /// </summary>
    /// <param name="source">源工作簿。</param>
    /// <param name="options">失败工作簿输出及清理选项。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="resolvedSheetRequests">实际工作表名称对应的导入请求。</param>
    internal static void ValidatePreflight(XLWorkbook source, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests)
    {
        if (options.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly)
            return;

        var candidateRows = errors.Where(error => error.RowIndex > 1)
            .Select(error => (error.SheetName, error.RowIndex))
            .Distinct(new ErrorRowComparer())
            .Count();
        if (options.MaxCandidateErrorRows.HasValue && candidateRows > options.MaxCandidateErrorRows.Value)
            throw ResourceLimit($"失败工作簿候选错误行超过限制: {options.MaxCandidateErrorRows.Value}",
                BingOfficesStage.Preflight);

        long cells = 0;
        long objects = 0;
        foreach (var group in errors.Where(error => error.RowIndex > 1)
                     .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            if (!source.Worksheets.TryGetWorksheet(group.Key, out var worksheet))
                continue;
            var headerRow = resolvedSheetRequests.TryGetValue(worksheet.Name, out var request)
                ? request.HeaderRowIndex + 1 : 1;
            var rows = new HashSet<int>(group.Select(error => error.RowIndex)) { headerRow };
            var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            cells += (long)rows.Count * lastColumn;
            objects += 1L + rows.Count + (long)rows.Count * lastColumn;
        }
        if (options.MaxCopiedCells.HasValue && cells > options.MaxCopiedCells.Value)
            throw ResourceLimit($"失败工作簿复制单元格估算数量超过限制: {options.MaxCopiedCells.Value}",
                BingOfficesStage.Preflight);
        if (options.MaxEstimatedTargetObjects.HasValue && objects > options.MaxEstimatedTargetObjects.Value)
            throw ResourceLimit($"失败工作簿目标对象估算数量超过限制: {options.MaxEstimatedTargetObjects.Value}",
                BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 创建仅包含表头和失败数据行的独立工作簿。
    /// </summary>
    /// <param name="source">源工作簿。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="resolvedSheetRequests">实际工作表名称对应的导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>仅包含表头和错误行的独立工作簿。</returns>
    private static XLWorkbook CreateErrorRowsWorkbook(XLWorkbook source,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken)
    {
        var output = new XLWorkbook();
        foreach (var group in errors.Where(error => error.RowIndex > 1)
                     .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!source.Worksheets.TryGetWorksheet(group.Key, out var sourceSheet))
                continue;
            var targetSheet = output.Worksheets.Add(sourceSheet.Name);
            var headerRow = resolvedSheetRequests.TryGetValue(sourceSheet.Name, out var request)
                ? request.HeaderRowIndex + 1 : 1;
            sourceSheet.Row(headerRow).CopyTo(targetSheet.Row(1));
            var lastColumn = sourceSheet.LastColumnUsed()?.ColumnNumber() ?? 0;
            for (var column = 1; column <= lastColumn; column++)
            {
                targetSheet.Column(column).Width = sourceSheet.Column(column).Width;
                if (sourceSheet.Column(column).IsHidden)
                    targetSheet.Column(column).Hide();
            }

            var targetRow = 2;
            var rowMap = new Dictionary<int, int>();
            foreach (var sourceRow in group.Select(error => error.RowIndex).Distinct().OrderBy(value => value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                sourceSheet.Row(sourceRow).CopyTo(targetSheet.Row(targetRow));
                rowMap[sourceRow] = targetRow++;
            }
            AddFailureColumns(targetSheet, group.ToArray(), rowMap, lastColumn + 1);
        }
        return output;
    }

    /// <summary>
    /// 将错误信息作为批注写入指定工作簿的单元格。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="conflictPolicy">已有批注的冲突处理策略。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void AnnotateErrors(XLWorkbook workbook, IEnumerable<ExcelImportError> errors,
        ExcelImportCommentConflictPolicy conflictPolicy, CancellationToken cancellationToken)
    {
        foreach (var error in errors.Where(error => error.RowIndex > 0 && error.ColumnIndex > 0))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!workbook.Worksheets.TryGetWorksheet(error.SheetName, out var worksheet))
                continue;
            var cell = worksheet.Cell(error.RowIndex, error.ColumnIndex);
            var message = error.Message ?? string.Empty;
            var author = "Bing.Offices";
            if (cell.HasComment)
            {
                if (conflictPolicy == ExcelImportCommentConflictPolicy.Preserve)
                    continue;
                if (conflictPolicy == ExcelImportCommentConflictPolicy.Fail)
                    throw new BingOfficesConfigurationException($"单元格已有失败批注目标: {cell.Address}",
                        stage: BingOfficesStage.Plan);
                var existing = cell.GetComment();
                if (conflictPolicy == ExcelImportCommentConflictPolicy.Append)
                {
                    message = existing + Environment.NewLine + message;
                    author = existing.Author;
                }
                existing.Delete();
            }
            var comment = cell.CreateComment();
            comment.Author = author;
            comment.AddText(message);
        }
    }

    /// <summary>
    /// 写入稳定的错误汇总工作表。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void WriteSummary(XLWorkbook workbook, IEnumerable<ExcelImportError> errors,
        CancellationToken cancellationToken)
    {
        var name = "_ImportErrors";
        var suffix = 1;
        while (workbook.Worksheets.Contains(name))
            name = $"_ImportErrors{suffix++}";
        var sheet = workbook.Worksheets.Add(name);
        var headers = new[] { "Code", "Message", "Sheet", "Row", "Column", "Property", "Header", "RawValue" };
        for (var column = 0; column < headers.Length; column++)
            sheet.Cell(1, column + 1).Value = headers[column];
        var row = 2;
        foreach (var error in errors)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Cell(row, 1).Value = error.Code.ToString();
            sheet.Cell(row, 2).Value = error.Message ?? string.Empty;
            sheet.Cell(row, 3).Value = error.SheetName ?? string.Empty;
            sheet.Cell(row, 4).Value = error.RowIndex;
            sheet.Cell(row, 5).Value = error.ColumnIndex;
            sheet.Cell(row, 6).Value = error.PropertyName ?? string.Empty;
            sheet.Cell(row, 7).Value = error.Header ?? string.Empty;
            sheet.Cell(row, 8).Value = FormatRawValue(error.RawValue);
            row++;
        }
    }

    /// <summary>
    /// 在错误行工作表末尾追加来源位置和错误摘要列。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="errors">结构化导入错误集合。</param>
    /// <param name="rowMap">源行号到失败工作簿行号的映射。</param>
    /// <param name="startColumn">错误信息起始列号，从 1 开始。</param>
    private static void AddFailureColumns(IXLWorksheet sheet, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<int, int> rowMap, int startColumn)
    {
        sheet.Cell(1, startColumn).Value = "__SourceSheet";
        sheet.Cell(1, startColumn + 1).Value = "__SourceRow";
        sheet.Cell(1, startColumn + 2).Value = "__ErrorCode";
        sheet.Cell(1, startColumn + 3).Value = "__Errors";
        foreach (var group in errors.GroupBy(error => error.RowIndex))
        {
            if (!rowMap.TryGetValue(group.Key, out var row))
                continue;
            sheet.Cell(row, startColumn).Value = group.First().SheetName ?? string.Empty;
            sheet.Cell(row, startColumn + 1).Value = group.Key;
            sheet.Cell(row, startColumn + 2).Value = string.Join(" | ", group.Select(error => error.Code));
            sheet.Cell(row, startColumn + 3).Value = string.Join(" | ", group.Select(error => error.Message)
                .Where(message => !string.IsNullOrWhiteSpace(message)));
        }
    }

    /// <summary>
    /// 将工作簿完整序列化到临时文件，并返回可提交产物。
    /// </summary>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="options">失败工作簿输出及清理选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>持有完整序列化内容的临时产物。</returns>
    private static ClosedXmlFailureWorkbookArtifact Serialize(XLWorkbook workbook,
        ExcelImportFailureOptions options, CancellationToken cancellationToken)
    {
        var directory = options.TemporaryDirectory ?? Path.GetTempPath();
        try
        {
            Directory.CreateDirectory(directory);
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            throw new BingOfficesImportException("失败工作簿临时目录创建失败。", exception, Provider,
                BingOfficesStage.Open);
        }
        var path = Path.Combine(directory, $"bing-offices-closedxml-failure-{Guid.NewGuid():N}.tmp");
        try
        {
            var stream = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                81920, FileOptions.SequentialScan);
            try
            {
                using (var limited = new LimitedWriteStream(stream, options.MaxSerializedBytes))
                    workbook.SaveAs(limited);
                cancellationToken.ThrowIfCancellationRequested();
                stream.Position = 0;
                return new ClosedXmlFailureWorkbookArtifact(stream, path, options);
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }
        catch (InvalidOperationException exception) when (exception.Message.StartsWith(
                   "失败工作簿超过最大序列化字节数:", StringComparison.Ordinal))
        {
            DeleteTemporaryFile(path, options, exception);
            throw ResourceLimit(exception.Message, BingOfficesStage.Serialize, exception);
        }
        catch (OperationCanceledException exception)
        {
            DeleteTemporaryFile(path, options, exception);
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            DeleteTemporaryFile(path, options, exception);
            throw new BingOfficesImportException("ClosedXML 失败工作簿序列化失败。", exception,
                Provider, BingOfficesStage.Serialize);
        }
    }

    /// <summary>
    /// 删除临时文件并隔离诊断接收器失败。
    /// </summary>
    /// <param name="path">待清理的临时文件路径。</param>
    /// <param name="options">失败工作簿输出及清理选项。</param>
    /// <param name="primaryException">需要附加清理失败信息的原始异常；无原始异常时为 null。</param>
    internal static void DeleteTemporaryFile(string path, ExcelImportFailureOptions options,
        Exception primaryException)
    {
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (Exception cleanupException) when (cleanupException is IOException
            || cleanupException is UnauthorizedAccessException)
        {
            var diagnostic = new ExcelImportFailureDiagnostic("FailureWorkbookTemporaryCleanupFailed",
                path, cleanupException);
            if (primaryException != null)
                primaryException.Data["Bing.Offices.FailureWorkbook.TemporaryCleanupException"] = cleanupException;
            try
            {
                options.DiagnosticSink?.Invoke(diagnostic);
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception diagnosticException) when (diagnosticException is not OutOfMemoryException
                && diagnosticException is not StackOverflowException)
            {
                Trace.WriteLine($"失败工作簿诊断接收器执行失败: {diagnosticException.GetType().Name}");
            }
        }
    }

    /// <summary>
    /// 格式化汇总表中的原始值。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>原始值的文本表示；null 转换为空字符串，二进制值转换为长度摘要。</returns>
    private static string FormatRawValue(object value) => value is byte[] bytes
        ? $"<binary:{bytes.Length}>"
        : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

    /// <summary>
    /// 创建失败工作簿资源限制异常。
    /// </summary>
    /// <param name="message">错误消息。</param>
    /// <param name="stage">发生错误的处理阶段。</param>
    /// <param name="innerException">引发资源限制的原始异常；没有时为 null。</param>
    /// <returns>包含处理阶段与原始异常的资源限制异常。</returns>
    private static BingOfficesResourceLimitException ResourceLimit(string message, BingOfficesStage stage,
        Exception innerException = null) => new(message, innerException, Provider,
        BingOfficesOperation.Import, stage);

    /// <summary>
    /// 限制序列化输出字节数的包装流。
    /// </summary>
    /// <remarks>释放包装流时仅刷新底层流，不关闭底层流。</remarks>
    private sealed class LimitedWriteStream : Stream
    {
        /// <summary>
        /// 接收序列化数据的底层流，由调用方负责释放。
        /// </summary>
        private readonly Stream _inner;
        /// <summary>
        /// 允许的最大序列化字节数；未指定时不限制。
        /// </summary>
        private readonly long? _maximum;

        /// <summary>
        /// 初始化一个 <see cref="LimitedWriteStream"/> 类型的实例。
        /// </summary>
        /// <param name="inner">接收序列化数据的底层流。</param>
        /// <param name="maximum">允许的最大序列化字节数；null 表示不限制。</param>
        internal LimitedWriteStream(Stream inner, long? maximum)
        {
            _inner = inner;
            _maximum = maximum;
        }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;
        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;
        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        /// <inheritdoc />
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value)
        {
            if (_maximum.HasValue && value > _maximum.Value)
                throw new InvalidOperationException($"失败工作簿超过最大序列化字节数: {_maximum.Value}");
            _inner.SetLength(value);
        }
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_maximum.HasValue && Position > _maximum.Value - count)
                throw new InvalidOperationException($"失败工作簿超过最大序列化字节数: {_maximum.Value}");
            _inner.Write(buffer, offset, count);
        }
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Flush();
            base.Dispose(disposing);
        }
    }
}

/// <summary>
/// 按 Sheet 名称（忽略大小写）和物理行号标识失败工作簿候选行。
/// </summary>
internal sealed class ErrorRowComparer : IEqualityComparer<(string SheetName, int RowIndex)>
{
    /// <inheritdoc />
    public bool Equals((string SheetName, int RowIndex) left, (string SheetName, int RowIndex) right) =>
        left.RowIndex == right.RowIndex
        && string.Equals(left.SheetName, right.SheetName, StringComparison.OrdinalIgnoreCase);

    /// <inheritdoc />
    public int GetHashCode((string SheetName, int RowIndex) value) =>
        HashCode.Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(value.SheetName ?? string.Empty), value.RowIndex);
}

/// <summary>
/// 表示已完成序列化、等待复制到调用方流的失败工作簿临时产物。
/// </summary>
internal sealed class ClosedXmlFailureWorkbookArtifact : IDisposable
{
    /// <summary>
    /// 保存完整失败工作簿的临时流，由当前产物负责释放。
    /// </summary>
    private readonly Stream _stream;
    /// <summary>
    /// 失败工作簿临时文件路径，用于释放产物时清理。
    /// </summary>
    private readonly string _path;
    /// <summary>
    /// 失败工作簿的临时文件清理与诊断选项。
    /// </summary>
    private readonly ExcelImportFailureOptions _options;
    /// <summary>
    /// 标记临时产物是否已释放，防止重复清理。
    /// </summary>
    private bool _disposed;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlFailureWorkbookArtifact"/> 类型的实例。
    /// </summary>
    /// <param name="stream">保存完整失败工作簿的临时流。</param>
    /// <param name="path">待清理的临时文件路径。</param>
    /// <param name="options">失败工作簿输出及清理选项。</param>
    internal ClosedXmlFailureWorkbookArtifact(Stream stream, string path, ExcelImportFailureOptions options)
    {
        _stream = stream;
        _path = path;
        _options = options;
    }

    /// <summary>
    /// 将失败工作簿复制到目标流。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal void CopyTo(Stream destination, CancellationToken cancellationToken)
    {
        ValidateDestination(destination);
        try
        {
            var buffer = new byte[81920];
            int count;
            while ((count = _stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                destination.Write(buffer, 0, count);
                cancellationToken.ThrowIfCancellationRequested();
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("失败工作簿复制到目标流失败。", exception,
                "ClosedXML", BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 异步将失败工作簿复制到目标流。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal async Task CopyToAsync(Stream destination, CancellationToken cancellationToken)
    {
        ValidateDestination(destination);
        try
        {
            await _stream.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("失败工作簿异步复制到目标流失败。", exception,
                "ClosedXML", BingOfficesStage.Write);
        }
    }

    /// <summary>
    /// 校验失败工作簿目标流可写。
    /// </summary>
    /// <param name="destination">接收失败工作簿内容的可写流。</param>
    private static void ValidateDestination(Stream destination)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
            return;
        _disposed = true;
        _stream.Dispose();
        ClosedXmlFailureWorkbookWriter.DeleteTemporaryFile(_path, _options, null);
    }
}
