using Bing.Offices.Extensions;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Imports;

/// <summary>Failure Workbook 错误行、样式、验证和图片复制职责。</summary>
internal static class NpoiFailureWorkbookCopier
{
    /// <summary>创建只包含错误数据行及相关元数据的独立工作簿。</summary>
    /// <param name="source">原始导入工作簿。</param>
    /// <param name="errors">用于筛选错误行的导入错误集合。</param>
    /// <param name="resolvedSheetRequests">实际工作表到请求的映射。</param>
    /// <param name="options">失败工作簿输出与诊断选项。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    /// <returns>与原始格式匹配的错误行工作簿。</returns>
    internal static IWorkbook CreateErrorRowsWorkbook(IWorkbook source, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        ExcelImportFailureOptions options, CancellationToken cancellationToken)
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
                    destination, styleCache, options);
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
    /// <param name="options">失败工作簿输出与诊断选项。</param>
    private static void CopyRow(IWorkbook source, IRow sourceRow, IRow destinationRow, IWorkbook destination,
        IDictionary<short, ICellStyle> styleCache, ExcelImportFailureOptions options)
    {
        if (sourceRow == null)
            return;
        destinationRow.Height = sourceRow.Height;
        NpoiFailureWorkbookDiagnostics.CopyOptionalRowMetadata(() => destinationRow.Hidden = sourceRow.Hidden,
            "Hidden", sourceRow.RowNum, options?.DiagnosticSink);
        NpoiFailureWorkbookDiagnostics.CopyOptionalRowMetadata(() => destinationRow.ZeroHeight = sourceRow.ZeroHeight,
            "ZeroHeight", sourceRow.RowNum, options?.DiagnosticSink);
        NpoiFailureWorkbookDiagnostics.CopyOptionalRowMetadata(() => destinationRow.Collapsed = sourceRow.Collapsed,
            "Collapsed", sourceRow.RowNum, options?.DiagnosticSink);
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
}

