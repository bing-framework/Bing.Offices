using Bing.Offices.Extensions;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// Failure Workbook 错误行、样式、验证和图片复制职责。
/// </summary>
internal static class NpoiFailureWorkbookCopier
{
    /// <summary>
    /// 创建只包含错误数据行及相关元数据的独立工作簿。
    /// </summary>
    /// <param name="source">原始导入工作簿。</param>
    /// <param name="errors">用于筛选错误行的导入错误集合。</param>
    /// <param name="resolvedSheetRequests">实际工作表到请求的映射。</param>
    /// <param name="options">失败工作簿输出与诊断选项。</param>
    /// <param name="cancellationToken">复制过程中检查的取消令牌。</param>
    /// <param name="requiredRowsBySheet">实体布局要求保留的零基结构行；通用 Workbook 为 null。</param>
    /// <returns>与原始格式匹配的错误行工作簿。</returns>
    internal static IWorkbook CreateErrorRowsWorkbook(IWorkbook source, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<string, ExcelSheetImportRequest> resolvedSheetRequests,
        ExcelImportFailureOptions options, CancellationToken cancellationToken,
        IReadOnlyDictionary<string, IReadOnlyList<int>> requiredRowsBySheet = null)
    {
        var destination = source is NPOI.HSSF.UserModel.HSSFWorkbook
            ? (IWorkbook)new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        var groups = errors.Where(error => error.RowIndex > 1)
            .GroupBy(error => error.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
        var rowMaps = new Dictionary<string, IReadOnlyDictionary<int, int>>(
            StringComparer.OrdinalIgnoreCase);
        var validationSheets = new List<(ISheet Source, ISheet Destination,
            IReadOnlyDictionary<int, int> RowMap)>();
        if (requiredRowsBySheet != null)
        {
            foreach (var sheetName in requiredRowsBySheet.Keys)
            {
                if (!groups.ContainsKey(sheetName))
                    groups.Add(sheetName, new List<ExcelImportError>());
            }
        }
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
            var sourceRows = new HashSet<int>(group.Value.Select(error => error.RowIndex - 1));
            if (group.Value.Count > 0)
                sourceRows.Add(headerRowIndex);
            if (requiredRowsBySheet != null
                && requiredRowsBySheet.TryGetValue(sourceSheet.SheetName, out var requiredRows))
            {
                sourceRows.UnionWith(requiredRows);
            }
            var rowMap = sourceRows.OrderBy(index => index)
                .Select((sourceRow, targetRow) => (sourceRow, targetRow))
                .ToDictionary(item => item.sourceRow, item => item.targetRow);
            rowMaps[sourceSheet.SheetName] = rowMap;
            var styleCache = new Dictionary<short, ICellStyle>();
            foreach (var pair in rowMap)
            {
                cancellationToken.ThrowIfCancellationRequested();
                CopyRow(source, sourceSheet.GetRow(pair.Key), destinationSheet.CreateRow(pair.Value),
                    destination, styleCache, options);
            }
            CopySheetMetadata(sourceSheet, destinationSheet, sourceRows);
            if (group.Value.Count > 0)
            {
                AddFailureColumns(destinationSheet, group.Value, rowMap,
                    sourceRows.Select(row => (int)(sourceSheet.GetRow(row)?.LastCellNum ?? 0))
                        .DefaultIfEmpty(0).Max());
            }
            CopyMergedRegions(sourceSheet, destinationSheet, rowMap);
            validationSheets.Add((sourceSheet, destinationSheet, rowMap));
            CopyPictures(sourceSheet, destinationSheet, rowMap, cancellationToken);
        }
        CopyDefinedNames(source, destination, rowMaps, cancellationToken);
        foreach (var item in validationSheets)
            CopyDataValidations(item.Source, item.Destination, item.RowMap, cancellationToken);
        return destination;
    }

    /// <summary>
    /// 复制能够完整映射到保留行的一维定义名称。
    /// </summary>
    /// <param name="source">稀疏源工作簿。</param>
    /// <param name="destination">错误行目标工作簿。</param>
    /// <param name="rowMaps">源工作表到目标行的映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void CopyDefinedNames(IWorkbook source, IWorkbook destination,
        IReadOnlyDictionary<string, IReadOnlyDictionary<int, int>> rowMaps,
        CancellationToken cancellationToken)
    {
        if (source is not HSSFWorkbook || destination == null || rowMaps == null
            || rowMaps.Count == 0)
            return;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sourceName in source.GetAllNames())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (sourceName == null || string.IsNullOrWhiteSpace(sourceName.NameName)
                || !TryParseDefinedNameFormula(sourceName.RefersToFormula, source,
                    sourceName.SheetIndex, out var sourceSheet, out var firstRow,
                    out var lastRow, out var firstColumn, out var lastColumn)
                || !rowMaps.TryGetValue(sourceSheet.SheetName, out var rowMap)
                || !TryMapDefinedNameRows(rowMap, firstRow, lastRow,
                    cancellationToken, out var outputFirstRow, out var outputLastRow))
                continue;

            var localSheetIndex = -1;
            if (sourceName.SheetIndex >= 0)
            {
                if (sourceName.SheetIndex >= source.NumberOfSheets)
                    continue;
                var localSheet = source.GetSheetAt(sourceName.SheetIndex);
                localSheetIndex = destination.GetSheetIndex(localSheet.SheetName);
                if (localSheetIndex < 0)
                    continue;
            }
            var key = sourceName.NameName + "\u001f" + localSheetIndex;
            if (!seen.Add(key))
                continue;
            var first = AbsoluteCellReference(outputFirstRow, firstColumn);
            var last = AbsoluteCellReference(outputLastRow, lastColumn);
            var destinationName = destination.CreateName();
            destinationName.NameName = sourceName.NameName;
            destinationName.SheetIndex = localSheetIndex;
            destinationName.RefersToFormula = QuoteDefinedNameSheet(sourceSheet.SheetName)
                + "!" + first;
            if (!string.Equals(first, last, StringComparison.Ordinal))
                destinationName.RefersToFormula += ":" + last;
        }
    }

    /// <summary>
    /// 解析单个一维绝对定义名称引用。
    /// </summary>
    /// <param name="formula">定义名称公式。</param>
    /// <param name="workbook">源工作簿。</param>
    /// <param name="currentSheetIndex">未限定工作表时使用的局部作用域。</param>
    /// <param name="rangeSheet">引用的源工作表。</param>
    /// <param name="firstRow">起始零基行号。</param>
    /// <param name="lastRow">结束零基行号。</param>
    /// <param name="firstColumn">起始零基列号。</param>
    /// <param name="lastColumn">结束零基列号。</param>
    /// <returns>公式受支持时返回 <see langword="true" />。</returns>
    private static bool TryParseDefinedNameFormula(string formula, IWorkbook workbook,
        int currentSheetIndex, out ISheet rangeSheet, out int firstRow, out int lastRow,
        out int firstColumn, out int lastColumn)
    {
        rangeSheet = null;
        firstRow = lastRow = firstColumn = lastColumn = 0;
        if (string.IsNullOrWhiteSpace(formula) || workbook == null)
            return false;
        var expression = formula.Trim();
        if (expression.StartsWith("=", StringComparison.Ordinal))
            expression = expression.Substring(1).Trim();
        var separator = -1;
        var quoted = false;
        for (var index = 0; index < expression.Length; index++)
        {
            var current = expression[index];
            if (current == '\'')
            {
                if (quoted && index + 1 < expression.Length && expression[index + 1] == '\'')
                    index++;
                else
                    quoted = !quoted;
            }
            else if (!quoted)
            {
                if (current == '!' && separator < 0)
                    separator = index;
                else if (current is '[' or ']' or ',' or '(' or ')' or ';' or '!')
                    return false;
            }
        }
        if (quoted)
            return false;
        if (separator >= 0)
        {
            var sheetToken = expression.Substring(0, separator).Trim();
            var sheetName = sheetToken;
            if (sheetToken.StartsWith("'", StringComparison.Ordinal))
            {
                if (sheetToken.Length < 2 || !sheetToken.EndsWith("'", StringComparison.Ordinal))
                    return false;
                sheetName = sheetToken.Substring(1, sheetToken.Length - 2).Replace("''", "'",
                    StringComparison.Ordinal);
            }
            if (string.IsNullOrWhiteSpace(sheetName))
                return false;
            rangeSheet = workbook.GetSheet(sheetName);
            if (rangeSheet == null)
                return false;
            expression = expression.Substring(separator + 1).Trim();
        }
        else
        {
            if (currentSheetIndex < 0 || currentSheetIndex >= workbook.NumberOfSheets)
                return false;
            rangeSheet = workbook.GetSheetAt(currentSheetIndex);
        }
        var parts = expression.Split(':');
        if (parts.Length < 1 || parts.Length > 2
            || !IsAbsoluteCellReference(parts[0].Trim())
            || (parts.Length == 2 && !IsAbsoluteCellReference(parts[1].Trim())))
            return false;
        try
        {
            var first = new CellReference(parts[0].Trim());
            var last = parts.Length == 1 ? first : new CellReference(parts[1].Trim());
            if (first.Row < 0 || first.Col < 0 || last.Row < first.Row || last.Col < first.Col
                || last.Row > workbook.SpreadsheetVersion.LastRowIndex
                || last.Col > workbook.SpreadsheetVersion.LastColumnIndex
                || (first.Row != last.Row && first.Col != last.Col))
                return false;
            firstRow = first.Row;
            lastRow = last.Row;
            firstColumn = first.Col;
            lastColumn = last.Col;
            return true;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// 判断地址是否为固定的 A1 单元格引用。
    /// </summary>
    /// <param name="address">待判断地址。</param>
    /// <returns>地址固定行列时返回 <see langword="true" />。</returns>
    private static bool IsAbsoluteCellReference(string address)
    {
        if (string.IsNullOrEmpty(address) || address[0] != '$')
            return false;
        var rowMarker = address.IndexOf('$', 1);
        if (rowMarker < 2 || rowMarker == address.Length - 1)
            return false;
        return address.Substring(1, rowMarker - 1).All(char.IsLetter)
            && address.Substring(rowMarker + 1).All(char.IsDigit)
            && address[rowMarker + 1] != '0';
    }

    /// <summary>
    /// 将定义名称的源行映射为目标行。
    /// </summary>
    /// <param name="rowMap">源行到目标行的映射。</param>
    /// <param name="firstRow">起始源行。</param>
    /// <param name="lastRow">结束源行。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="firstOutputRow">起始目标行。</param>
    /// <param name="lastOutputRow">结束目标行。</param>
    /// <returns>区域完整保留且连续时返回 <see langword="true" />。</returns>
    private static bool TryMapDefinedNameRows(IReadOnlyDictionary<int, int> rowMap,
        int firstRow, int lastRow, CancellationToken cancellationToken,
        out int firstOutputRow, out int lastOutputRow)
    {
        firstOutputRow = lastOutputRow = 0;
        if (rowMap == null || firstRow < 0 || lastRow < firstRow)
            return false;
        if (!rowMap.TryGetValue(firstRow, out firstOutputRow))
            return false;
        if (firstRow == lastRow)
        {
            lastOutputRow = firstOutputRow;
            return true;
        }
        for (var sourceRow = firstRow + 1; sourceRow <= lastRow; sourceRow++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!rowMap.TryGetValue(sourceRow, out var outputRow)
                || outputRow != firstOutputRow + sourceRow - firstRow)
                return false;
            lastOutputRow = outputRow;
        }
        return true;
    }

    /// <summary>
    /// 创建定义名称公式中的绝对 A1 地址。
    /// </summary>
    /// <param name="row">零基行号。</param>
    /// <param name="column">零基列号。</param>
    /// <returns>绝对 A1 地址。</returns>
    private static string AbsoluteCellReference(int row, int column)
        => new CellReference(row, column, true, true).FormatAsString();

    /// <summary>
    /// 将工作表名称编码为定义名称公式前缀。
    /// </summary>
    /// <param name="sheetName">工作表名称。</param>
    /// <returns>带单引号的工作表名称。</returns>
    private static string QuoteDefinedNameSheet(string sheetName)
        => "'" + (sheetName ?? string.Empty).Replace("'", "''", StringComparison.Ordinal) + "'";

    /// <summary>
    /// 复制行高、隐藏状态、单元格值、样式、超链接和批注。
    /// </summary>
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

    /// <summary>
    /// 复制富文本内容及其格式运行到目标工作簿。
    /// </summary>
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

    /// <summary>
    /// 将源样式复制到目标工作簿并按样式索引复用。
    /// </summary>
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

    /// <summary>
    /// 复制单元格超链接并重新绑定目标单元格坐标。
    /// </summary>
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

    /// <summary>
    /// 复制单元格批注文本、作者、可见性和相对锚点。
    /// </summary>
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

    /// <summary>
    /// 复制工作表显示属性、列宽、列隐藏状态和冻结窗格。
    /// </summary>
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

    /// <summary>
    /// 在错误行工作表末尾添加来源位置和错误摘要列。
    /// </summary>
    /// <param name="sheet">目标错误行工作表。</param>
    /// <param name="errors">当前工作表的错误集合。</param>
    /// <param name="rowMap">源零基行号到目标零基行号的映射。</param>
    /// <param name="sourceColumn">源内容之后的起始列索引。</param>
    private static void AddFailureColumns(ISheet sheet, IReadOnlyCollection<ExcelImportError> errors,
        IReadOnlyDictionary<int, int> rowMap, int sourceColumn)
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

    /// <summary>
    /// 复制完全落在错误行集合中的合并区域。
    /// </summary>
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

    /// <summary>
    /// 复制完全落在错误行集合中的工作簿原生数据校验。
    /// </summary>
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
                try
                {
                    destination.AddValidationData(copied);
                }
                catch (FormulaParseException)
                {
                    // 命名列表来源未随错误行复制时，跳过无法解析的规则，避免破坏失败工作簿输出。
                }
            }
        }
    }

    /// <summary>
    /// 复制锚点行完全存在于错误行集合中的图片资源。
    /// </summary>
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
