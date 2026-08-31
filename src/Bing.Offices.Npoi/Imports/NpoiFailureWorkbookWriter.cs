using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 生成并写出导入失败工作簿，隔离失败产物的复制、注释和摘要逻辑。
/// </summary>
internal static class NpoiFailureWorkbookWriter
{
    /// <summary>
    /// 根据失败选项写出失败工作簿；没有配置失败产物或没有错误时不执行任何操作。
    /// </summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Write(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken)
        => Write(workbook, options, errors, resolvedSheetRequests, cancellationToken,
            new SystemFailureWorkbookFileSystem());

    /// <summary>使用指定文件系统适配器写出失败工作簿，便于测试临时文件操作。</summary>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="resolvedSheetRequests">实际解析 Sheet 名称到请求的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileSystem">临时文件系统适配器。</param>
    internal static void Write(IWorkbook workbook, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken, IFailureWorkbookFileSystem fileSystem)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        if (fileSystem == null)
            throw new ArgumentNullException(nameof(fileSystem));
        cancellationToken.ThrowIfCancellationRequested();
        IWorkbook outputWorkbook = workbook;
        IWorkbook independentWorkbook = null;
        try
        {
            if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly)
                outputWorkbook = independentWorkbook = CreateErrorRowsWorkbook(workbook, errors, resolvedSheetRequests,
                    cancellationToken);
            else
                AnnotateErrors(outputWorkbook, errors, options.CommentConflictPolicy);

            WriteFailureSummary(outputWorkbook, errors, cancellationToken);
            var temporaryDirectory = options.TemporaryDirectory ?? Path.GetTempPath();
            try
            {
                fileSystem.CreateDirectory(temporaryDirectory);
            }
            catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
            {
                throw new IOException("失败工作簿临时目录创建失败。", exception);
            }
            var temporaryPath = Path.Combine(temporaryDirectory, $"bing-offices-failure-{Guid.NewGuid():N}.tmp");
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                Stream output;
                try
                {
                    output = fileSystem.CreateFile(temporaryPath);
                }
                catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
                {
                    throw new IOException("失败工作簿临时文件创建失败。", exception);
                }
                using (output)
                using (var limitedOutput = new LimitedWriteStream(output, options.MaxBytes))
                {
                    try
                    {
                        outputWorkbook.Write(limitedOutput, false);
                    }
                    catch (Exception exception)
                    {
                        var limitException = FindLimitException(exception);
                        if (limitException != null)
                            throw limitException;
                        throw new InvalidOperationException("失败工作簿序列化失败。", exception);
                    }
                    limitedOutput.Flush();
                    if (options.MaxBytes.HasValue && output.Length > options.MaxBytes.Value)
                        throw new InvalidOperationException($"失败工作簿超过最大字节数: {options.MaxBytes.Value}");
                    cancellationToken.ThrowIfCancellationRequested();
                    output.Position = 0;
                    try
                    {
                        WriteStream(options.Destination, output, cancellationToken);
                    }
                    catch (Exception exception) when (!(exception is OperationCanceledException))
                    {
                        throw new InvalidOperationException("失败工作簿复制到目标流失败。", exception);
                    }
                }
            }
            catch (Exception exception)
            {
                DeleteTemporaryFile(options, temporaryPath, exception, fileSystem);
                throw;
            }
            DeleteTemporaryFile(options, temporaryPath, null, fileSystem);
        }
        finally
        {
            independentWorkbook?.Close();
        }
    }

    /// <summary>删除失败工作簿临时文件，并将清理失败写入诊断或主异常。</summary>
    /// <param name="options">失败工作簿输出选项和诊断接收器。</param>
    /// <param name="temporaryPath">待删除的临时文件路径。</param>
    /// <param name="primaryException">提交失败时的主异常。</param>
    /// <param name="fileSystem">临时文件系统适配器。</param>
    private static void DeleteTemporaryFile(ExcelImportFailureOptions options, string temporaryPath,
        Exception primaryException, IFailureWorkbookFileSystem fileSystem)
    {
        try
        {
            fileSystem.Delete(temporaryPath);
        }
        catch (Exception cleanupException) when (cleanupException is IOException
            || cleanupException is UnauthorizedAccessException)
        {
            var diagnostic = new ExcelImportFailureDiagnostic("FailureWorkbookTemporaryCleanupFailed",
                temporaryPath, cleanupException);
            if (options.DiagnosticSink != null)
            {
                try
                {
                    options.DiagnosticSink(diagnostic);
                }
                catch (Exception diagnosticException)
                {
                    Trace.WriteLine($"失败工作簿诊断接收器执行失败: {diagnosticException.GetType().Name}");
                }
            }
            else if (primaryException != null)
            {
                primaryException.Data["Bing.Offices.FailureWorkbook.TemporaryCleanupException"] = cleanupException;
            }
            else
            {
                throw new IOException("失败工作簿临时文件清理失败。", cleanupException);
            }
        }
    }

    /// <summary>在失败工作簿中写入错误汇总工作表。</summary>
    /// <param name="workbook">待写入汇总的工作簿。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="cancellationToken">写入过程中检查的取消令牌。</param>
    private static void WriteFailureSummary(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
        CancellationToken cancellationToken)
    {
        var summaryName = "_ImportErrors";
        var suffix = 1;
        while (workbook.GetSheet(summaryName) != null)
            summaryName = $"_ImportErrors{suffix++}";
        var summary = workbook.CreateSheet(summaryName);
        var header = summary.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Message");
        header.CreateCell(2).SetCellValue("Sheet");
        header.CreateCell(3).SetCellValue("Row");
        header.CreateCell(4).SetCellValue("Column");
        header.CreateCell(5).SetCellValue("Property");
        header.CreateCell(6).SetCellValue("Header");
        header.CreateCell(7).SetCellValue("RawValue");
        var rowIndex = 1;
        foreach (var error in errors)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = summary.CreateRow(rowIndex++);
            row.CreateCell(0).SetCellValue(error.Code.ToString());
            row.CreateCell(1).SetCellValue(error.Message ?? string.Empty);
            row.CreateCell(2).SetCellValue(error.SheetName ?? string.Empty);
            row.CreateCell(3).SetCellValue(error.RowIndex);
            row.CreateCell(4).SetCellValue(error.ColumnIndex);
            row.CreateCell(5).SetCellValue(error.PropertyName ?? string.Empty);
            row.CreateCell(6).SetCellValue(error.Header ?? string.Empty);
            row.CreateCell(7).SetCellValue(FormatRawValue(error.RawValue));
        }
    }

    /// <summary>创建只包含错误数据行及相关元数据的独立工作簿。</summary>
    /// <param name="source">原始导入工作簿。</param>
    /// <param name="errors">用于筛选错误行的导入错误集合。</param>
    /// <param name="resolvedSheetRequests">实际工作表到请求的映射。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    /// <returns>与原始格式匹配的错误行工作簿。</returns>
    private static IWorkbook CreateErrorRowsWorkbook(IWorkbook source, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        CancellationToken cancellationToken)
    {
        var destination = source is NPOI.HSSF.UserModel.HSSFWorkbook
            ? (IWorkbook)new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        var groups = errors.Where(error => error.RowIndex > 1)
            .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        foreach (var group in groups)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceSheet = source.GetSheet(group.Key);
            if (sourceSheet == null)
                continue;
            var destinationSheet = destination.CreateSheet(sourceSheet.SheetName);
            var headerRowIndex = resolvedSheetRequests.TryGetValue(sourceSheet.SheetName, out var request)
                ? request.HeaderRowIndex
                : sourceSheet.FirstRowNum;
            var sourceRows = new HashSet<int>(group.Value.Select(error => error.RowIndex - 1))
            {
                headerRowIndex
            };
            var rowMap = sourceRows.OrderBy(index => index)
                .Select((sourceRow, targetRow) => (sourceRow, targetRow))
                .ToDictionary(item => item.sourceRow, item => item.targetRow);
            var styleCache = new Dictionary<short, ICellStyle>();
            foreach (var pair in rowMap)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CopyRow(source, sourceSheet.GetRow(pair.Key), destinationSheet.CreateRow(pair.Value),
                    destination, styleCache);
            }
            CopySheetMetadata(sourceSheet, destinationSheet, sourceRows);
            AddFailureColumns(destinationSheet, group.Value, rowMap,
                sourceSheet.GetRow(headerRowIndex)?.LastCellNum ?? 0);
            CopyMergedRegions(sourceSheet, destinationSheet, rowMap);
            CopyDataValidations(sourceSheet, destinationSheet, rowMap, cancellationToken);
            CopyPictures(sourceSheet, destinationSheet, rowMap, cancellationToken);
        }
        return destination;
    }

    /// <summary>复制行高、隐藏状态、单元格值、样式、超链接和批注。</summary>
    /// <param name="source">源工作簿。</param>
    /// <param name="sourceRow">源数据行。</param>
    /// <param name="destinationRow">目标数据行。</param>
    /// <param name="destination">目标工作簿。</param>
    /// <param name="styleCache">按源样式索引缓存目标样式的字典。</param>
    private static void CopyRow(IWorkbook source, IRow sourceRow, IRow destinationRow, IWorkbook destination,
        IDictionary<short, ICellStyle> styleCache)
    {
        if (sourceRow == null)
            return;
        destinationRow.Height = sourceRow.Height;
        try
        {
            destinationRow.Hidden = sourceRow.Hidden;
        }
        catch (NotImplementedException)
        {
        }
        try
        {
            destinationRow.ZeroHeight = sourceRow.ZeroHeight;
        }
        catch (NotImplementedException)
        {
        }
        try
        {
            destinationRow.Collapsed = sourceRow.Collapsed;
        }
        catch (NotImplementedException)
        {
        }
        if (sourceRow.RowStyle != null)
            destinationRow.RowStyle = CloneStyle(sourceRow.RowStyle, destination, styleCache);
        foreach (var sourceCell in sourceRow.Cells)
        {
            var destinationCell = destinationRow.CreateCell(sourceCell.ColumnIndex);
            switch (sourceCell.CellType)
            {
                case CellType.Formula:
                    destinationCell.SetCellFormula(sourceCell.CellFormula);
                    break;
                case CellType.Numeric:
                    destinationCell.SetCellValue(sourceCell.NumericCellValue);
                    break;
                case CellType.Boolean:
                    destinationCell.SetCellValue(sourceCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    destinationCell.SetCellErrorValue(sourceCell.ErrorCellValue);
                    break;
                case CellType.String:
                    destinationCell.SetCellValue(CopyRichText(sourceCell.RichStringCellValue, source, destination));
                    break;
            }
            if (sourceCell.CellStyle != null)
            {
                destinationCell.CellStyle = CloneStyle(sourceCell.CellStyle, destination, styleCache);
            }
            CopyHyperlink(sourceCell, destinationCell, destination);
            CopyComment(sourceCell, destinationCell, destination);
        }
    }

    /// <summary>复制富文本内容及其格式运行到目标工作簿。</summary>
    /// <param name="source">源富文本。</param>
    /// <param name="sourceWorkbook">源工作簿，用于 HSSF 字体解析。</param>
    /// <param name="destination">目标工作簿。</param>
    /// <returns>目标工作簿中的富文本副本。</returns>
    private static IRichTextString CopyRichText(IRichTextString source, IWorkbook sourceWorkbook,
        IWorkbook destination)
    {
        if (source == null)
            return destination.GetCreationHelper().CreateRichTextString(string.Empty);
        var copied = destination.GetCreationHelper().CreateRichTextString(source.String ?? string.Empty);
        if (source is XSSFRichTextString xssfSource && copied is XSSFRichTextString xssfCopied)
        {
            for (var run = 0; run < xssfSource.NumFormattingRuns; run++)
            {
                var start = xssfSource.GetIndexOfFormattingRun(run);
                var end = start + xssfSource.GetLengthOfFormattingRun(run);
                var font = destination.CreateFont();
                font.CloneStyleFrom(xssfSource.GetFontOfFormattingRun(run));
                xssfCopied.ApplyFont(start, end, font);
            }
        }
        else if (source is HSSFRichTextString hssfSource && copied is HSSFRichTextString hssfCopied)
        {
            for (var run = 0; run < hssfSource.NumFormattingRuns; run++)
            {
                var start = hssfSource.GetIndexOfFormattingRun(run);
                var end = run + 1 < hssfSource.NumFormattingRuns
                    ? hssfSource.GetIndexOfFormattingRun(run + 1)
                    : hssfSource.String.Length;
                var font = destination.CreateFont();
                font.CloneStyleFrom(((HSSFWorkbook)sourceWorkbook).GetFontAt(
                    hssfSource.GetFontOfFormattingRun(run)));
                hssfCopied.ApplyFont(start, end, font.Index);
            }
        }
        return copied;
    }

    /// <summary>将源样式复制到目标工作簿并按样式索引复用。</summary>
    /// <param name="sourceStyle">源工作簿样式。</param>
    /// <param name="destination">目标工作簿。</param>
    /// <param name="styleCache">按源样式索引缓存目标样式的字典。</param>
    /// <returns>目标工作簿中的样式副本。</returns>
    private static ICellStyle CloneStyle(ICellStyle sourceStyle, IWorkbook destination,
        IDictionary<short, ICellStyle> styleCache)
    {
        if (!styleCache.TryGetValue(sourceStyle.Index, out var style))
        {
            style = destination.CreateCellStyle();
            style.CloneStyleFrom(sourceStyle);
            styleCache[sourceStyle.Index] = style;
        }
        return style;
    }

    /// <summary>复制单元格超链接并重新绑定目标单元格坐标。</summary>
    /// <param name="sourceCell">源单元格。</param>
    /// <param name="destinationCell">目标单元格。</param>
    /// <param name="destination">目标工作簿。</param>
    private static void CopyHyperlink(ICell sourceCell, ICell destinationCell, IWorkbook destination)
    {
        var sourceHyperlink = sourceCell.Hyperlink;
        if (sourceHyperlink == null)
            return;
        var hyperlink = destination.GetCreationHelper().CreateHyperlink(sourceHyperlink.Type);
        hyperlink.Address = sourceHyperlink.Address;
        hyperlink.Label = sourceHyperlink.Label;
        hyperlink.FirstRow = destinationCell.RowIndex;
        hyperlink.LastRow = destinationCell.RowIndex;
        hyperlink.FirstColumn = destinationCell.ColumnIndex;
        hyperlink.LastColumn = destinationCell.ColumnIndex;
        destinationCell.Hyperlink = hyperlink;
    }

    /// <summary>复制单元格批注文本、作者、可见性和相对锚点。</summary>
    /// <param name="sourceCell">源单元格。</param>
    /// <param name="destinationCell">目标单元格。</param>
    /// <param name="destination">目标工作簿。</param>
    private static void CopyComment(ICell sourceCell, ICell destinationCell, IWorkbook destination)
    {
        var sourceComment = sourceCell.CellComment;
        if (sourceComment == null)
            return;
        var anchor = destination.GetCreationHelper().CreateClientAnchor();
        var sourceAnchor = sourceComment.ClientAnchor;
        anchor.Col1 = destinationCell.ColumnIndex;
        anchor.Col2 = destinationCell.ColumnIndex + Math.Max(1, sourceAnchor?.Col2 - sourceAnchor.Col1 ?? 2);
        anchor.Row1 = destinationCell.RowIndex;
        anchor.Row2 = destinationCell.RowIndex + Math.Max(1, sourceAnchor?.Row2 - sourceAnchor.Row1 ?? 3);
        var comment = destinationCell.Sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
        comment.String = destination.GetCreationHelper().CreateRichTextString(sourceComment.String?.String ?? string.Empty);
        comment.Author = sourceComment.Author;
        comment.Visible = sourceComment.Visible;
        destinationCell.CellComment = comment;
    }

    /// <summary>复制工作表显示属性、列宽、列隐藏状态和冻结窗格。</summary>
    /// <param name="source">源工作表。</param>
    /// <param name="destination">目标工作表。</param>
    /// <param name="sourceRows">已复制到目标工作表的源行集合。</param>
    private static void CopySheetMetadata(ISheet source, ISheet destination, ISet<int> sourceRows)
    {
        destination.DefaultColumnWidth = source.DefaultColumnWidth;
        destination.DefaultRowHeight = source.DefaultRowHeight;
        destination.DefaultRowHeightInPoints = source.DefaultRowHeightInPoints;
        destination.DisplayGridlines = source.DisplayGridlines;
        destination.DisplayFormulas = source.DisplayFormulas;
        destination.DisplayZeros = source.DisplayZeros;
        destination.DisplayRowColHeadings = source.DisplayRowColHeadings;
        destination.IsRightToLeft = source.IsRightToLeft;
        destination.HorizontallyCenter = source.HorizontallyCenter;
        destination.VerticallyCenter = source.VerticallyCenter;
        destination.FitToPage = source.FitToPage;
        destination.ForceFormulaRecalculation = source.ForceFormulaRecalculation;

        var maxColumn = Enumerable.Range(source.FirstRowNum, Math.Max(0, source.LastRowNum - source.FirstRowNum + 1))
            .Select(rowIndex => (int)(source.GetRow(rowIndex)?.LastCellNum ?? 0))
            .DefaultIfEmpty(0).Max();
        for (var column = 0; column < maxColumn; column++)
        {
            destination.SetColumnWidth(column, source.GetColumnWidth(column));
            destination.SetColumnHidden(column, source.IsColumnHidden(column));
        }

        var pane = source.PaneInformation;
        if (pane != null && pane.IsFreezePane())
        {
            var splitRow = sourceRows.Count(row => row < pane.VerticalSplitPosition);
            destination.CreateFreezePane(pane.HorizontalSplitPosition, splitRow);
        }
    }

    /// <summary>在错误行工作表末尾添加来源位置和错误摘要列。</summary>
    /// <param name="sheet">目标错误行工作表。</param>
    /// <param name="errors">当前工作表的错误集合。</param>
    /// <param name="rowMap">源零基行号到目标零基行号的映射。</param>
    /// <param name="sourceColumn">源表头之后的起始列索引。</param>
    private static void AddFailureColumns(ISheet sheet, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<int, int> rowMap, short sourceColumn)
    {
        var startColumn = Math.Max(0, (int)sourceColumn);
        var header = sheet.GetRow(0) ?? sheet.CreateRow(0);
        header.CreateCell(startColumn).SetCellValue("__SourceSheet");
        header.CreateCell(startColumn + 1).SetCellValue("__SourceRow");
        header.CreateCell(startColumn + 2).SetCellValue("__ErrorCode");
        header.CreateCell(startColumn + 3).SetCellValue("__Errors");
        foreach (var group in errors.Where(error => error.RowIndex > 1).GroupBy(error => error.RowIndex - 1))
        {
            if (!rowMap.TryGetValue(group.Key, out var targetRow))
                continue;
            var row = sheet.GetRow(targetRow) ?? sheet.CreateRow(targetRow);
            row.CreateCell(startColumn).SetCellValue(group.First().SheetName ?? string.Empty);
            row.CreateCell(startColumn + 1).SetCellValue(group.Key + 1);
            row.CreateCell(startColumn + 2).SetCellValue(string.Join(" | ", group.Select(error => error.Code)));
            row.CreateCell(startColumn + 3).SetCellValue(string.Join(" | ", group.Select(error => error.Message)
                .Where(message => !string.IsNullOrWhiteSpace(message))));
        }
    }

    /// <summary>复制完全落在错误行集合中的合并区域。</summary>
    /// <param name="source">源工作表。</param>
    /// <param name="destination">目标工作表。</param>
    /// <param name="rowMap">源零基行号到目标零基行号的映射。</param>
    private static void CopyMergedRegions(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap)
    {
        foreach (var region in source.MergedRegions)
        {
            if (!rowMap.TryGetValue(region.FirstRow, out var firstRow)
                || !rowMap.TryGetValue(region.LastRow, out var lastRow))
                continue;
            var contiguous = true;
            for (var row = region.FirstRow; row <= region.LastRow; row++)
                contiguous &= rowMap.ContainsKey(row);
            if (contiguous)
                destination.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(firstRow, lastRow,
                    region.FirstColumn, region.LastColumn));
        }
    }

    /// <summary>复制完全落在错误行集合中的工作簿原生数据校验。</summary>
    /// <param name="source">源工作表。</param>
    /// <param name="destination">目标工作表。</param>
    /// <param name="rowMap">源零基行号到目标零基行号的映射。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    private static void CopyDataValidations(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap, CancellationToken cancellationToken)
    {
        var helper = destination.GetDataValidationHelper();
        foreach (var validation in source.GetDataValidations())
        {
            foreach (var region in validation.Regions.CellRangeAddresses)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!rowMap.TryGetValue(region.FirstRow, out var firstRow)
                    || !rowMap.TryGetValue(region.LastRow, out var lastRow))
                    continue;
                var contiguous = true;
                for (var row = region.FirstRow; row <= region.LastRow; row++)
                    contiguous &= rowMap.ContainsKey(row);
                if (!contiguous)
                    continue;
                var targetRegion = new NPOI.SS.Util.CellRangeAddressList(firstRow, lastRow,
                    region.FirstColumn, region.LastColumn);
                var copied = helper.CreateValidation(validation.ValidationConstraint, targetRegion);
                copied.EmptyCellAllowed = validation.EmptyCellAllowed;
                copied.ShowErrorBox = validation.ShowErrorBox;
                if (validation.ShowErrorBox)
                    copied.CreateErrorBox(validation.ErrorBoxTitle, validation.ErrorBoxText);
                copied.ShowPromptBox = validation.ShowPromptBox;
                if (validation.ShowPromptBox)
                    copied.CreatePromptBox(validation.PromptBoxTitle, validation.PromptBoxText);
                copied.SuppressDropDownArrow = validation.SuppressDropDownArrow;
                copied.ErrorStyle = validation.ErrorStyle;
                destination.AddValidationData(copied);
            }
        }
    }

    /// <summary>复制锚点行完全存在于错误行集合中的图片资源。</summary>
    /// <param name="source">源工作表。</param>
    /// <param name="destination">目标工作表。</param>
    /// <param name="rowMap">源零基行号到目标零基行号的映射。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    private static void CopyPictures(ISheet source, ISheet destination,
        IReadOnlyDictionary<int, int> rowMap, CancellationToken cancellationToken)
    {
        foreach (var picture in source.GetAllPictureInfos())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sourceLastRow = Math.Max(picture.MinRow, picture.MaxRow - 1);
            if (!rowMap.TryGetValue(picture.MinRow, out var minRow)
                || !rowMap.TryGetValue(sourceLastRow, out var mappedLastRow))
                continue;
            var contiguous = true;
            for (var row = picture.MinRow; row <= sourceLastRow; row++)
                contiguous &= rowMap.ContainsKey(row);
            if (!contiguous)
                continue;
            destination.AddPicture(new PictureInfo(minRow, mappedLastRow + 1, picture.MinCol, picture.MaxCol,
                picture.PictureData, picture.PictureStyle ?? new PictureStyle()));
        }
    }

    /// <summary>将错误原始值格式化为汇总工作表可写入的文本。</summary>
    /// <param name="value">错误原始值。</param>
    /// <returns>二进制值的长度标记或普通区域性无关文本。</returns>
    private static string FormatRawValue(object value)
    {
        if (value is byte[] bytes)
            return $"<binary:{bytes.Length}>";
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    /// <summary>按冲突策略将错误消息写入原始工作簿的目标单元格批注。</summary>
    /// <param name="workbook">待添加批注的工作簿。</param>
    /// <param name="errors">已收集的导入错误。</param>
    /// <param name="conflictPolicy">已有批注时的处理策略。</param>
    private static void AnnotateErrors(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
        ExcelImportCommentConflictPolicy conflictPolicy)
    {
        foreach (var error in errors.Where(error => error.RowIndex > 0 && error.ColumnIndex > 0))
        {
            var sheet = workbook.GetSheet(error.SheetName);
            var cell = sheet?.GetRow(error.RowIndex - 1)?.GetCell(error.ColumnIndex - 1);
            if (cell == null)
                continue;
            var existing = cell.CellComment;
            if (existing != null && conflictPolicy == ExcelImportCommentConflictPolicy.Preserve)
                continue;
            if (existing != null && conflictPolicy == ExcelImportCommentConflictPolicy.Fail)
                throw new InvalidOperationException($"单元格已有失败批注目标: {cell.Address}");
            var text = error.Message ?? string.Empty;
            if (existing != null && conflictPolicy == ExcelImportCommentConflictPolicy.Append)
                text = existing.String.String + Environment.NewLine + text;
            if (existing != null)
            {
                existing.String = workbook.GetCreationHelper().CreateRichTextString(text);
                if (conflictPolicy == ExcelImportCommentConflictPolicy.Replace)
                    existing.Author = "Bing.Offices";
                continue;
            }
            var anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = cell.ColumnIndex + 2;
            anchor.Row1 = cell.RowIndex;
            anchor.Row2 = cell.RowIndex + 3;
            var comment = sheet.CreateDrawingPatriarch().CreateCellComment(anchor);
            comment.String = workbook.GetCreationHelper().CreateRichTextString(text);
            comment.Author = "Bing.Offices";
            cell.CellComment = comment;
        }
    }

    /// <summary>将临时工作簿流复制到调用方目标流，并在块边界检查取消。</summary>
    /// <param name="destination">调用方提供的可写目标流。</param>
    /// <param name="source">临时工作簿源流。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    private static void WriteStream(Stream destination, Stream source, CancellationToken cancellationToken)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        var buffer = new byte[81920];
        int count;
        while ((count = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, count);
            cancellationToken.ThrowIfCancellationRequested();
        }
    }

    /// <summary>
    /// 查找 NPOI 序列化层包装的大小限制异常。
    /// </summary>
    /// <param name="exception">待检查的异常。</param>
    /// <returns>找到的大小限制异常；否则返回 null。</returns>
    private static InvalidOperationException FindLimitException(Exception exception)
    {
        while (exception != null)
        {
            if (exception is InvalidOperationException invalidOperationException
                && invalidOperationException.Message.StartsWith("失败工作簿超过最大字节数:",
                    StringComparison.Ordinal))
                return invalidOperationException;
            exception = exception.InnerException;
        }
        return null;
    }

    /// <summary>
    /// 在序列化阶段限制失败工作簿的最大字节数。
    /// </summary>
    private sealed class LimitedWriteStream : Stream
    {
        /// <summary>由调用方拥有且仅由包装器刷新、不负责释放的底层输出流。</summary>
        private readonly Stream _inner;
        /// <summary>失败工作簿序列化允许写入的最大字节数。</summary>
        private readonly long? _maxBytes;

        /// <summary>
        /// 初始化受限写入流。
        /// </summary>
        /// <param name="inner">实际写入流。</param>
        /// <param name="maxBytes">最大允许字节数。</param>
        internal LimitedWriteStream(Stream inner, long? maxBytes)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _maxBytes = maxBytes;
        }

        /// <inheritdoc />
        public override bool CanRead => false;

        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;

        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;

        /// <inheritdoc />
        public override long Length => _inner.Length;

        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => _inner.Position = value;
        }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        /// <inheritdoc />
        public override void SetLength(long value)
        {
            if (_maxBytes.HasValue && value > _maxBytes.Value)
                throw new InvalidOperationException($"失败工作簿超过最大字节数: {_maxBytes.Value}");
            _inner.SetLength(value);
        }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_maxBytes.HasValue && Position > _maxBytes.Value - count)
                throw new InvalidOperationException($"失败工作簿超过最大字节数: {_maxBytes.Value}");
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

/// <summary>为失败工作簿临时文件创建和清理抽象的文件系统操作。</summary>
internal interface IFailureWorkbookFileSystem
{
    /// <summary>确保临时输出目录存在。</summary>
    void CreateDirectory(string path);
    /// <summary>以读写方式独占创建临时工作簿文件。</summary>
    Stream CreateFile(string path);
    /// <summary>删除不再需要的临时工作簿文件。</summary>
    void Delete(string path);
}

/// <summary>基于 <see cref="Directory"/> 和 <see cref="File"/> 的失败工作簿文件系统实现。</summary>
internal sealed class SystemFailureWorkbookFileSystem : IFailureWorkbookFileSystem
{
    /// <inheritdoc />
    public void CreateDirectory(string path) => Directory.CreateDirectory(path);

    /// <inheritdoc />
    public Stream CreateFile(string path) => new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
        FileShare.None, 81920, FileOptions.SequentialScan);

    /// <inheritdoc />
    public void Delete(string path) => File.Delete(path);
}
