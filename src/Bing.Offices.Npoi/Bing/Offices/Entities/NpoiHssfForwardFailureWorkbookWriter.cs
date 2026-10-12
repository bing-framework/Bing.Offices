using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 写出 HSSF Forward 的低保留错误行工作簿。
/// </summary>
/// <remarks>
/// 先由捕获快照构造只包含表头和失败行的稀疏 HSSF 源，再复用 NPOI 失败工作簿写出器，
/// 从而保持 XLS 格式、资源限制、文件提交和目标流所有权契约。
/// </remarks>
internal static class NpoiHssfForwardFailureWorkbookWriter
{
    private static readonly IReadOnlyDictionary<string, ExcelSheetImportRequest>
        EmptyResolvedSheetRequests = new Dictionary<string, ExcelSheetImportRequest>(
            StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 判断捕获快照是否覆盖所有需要写出的数据错误行。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="errors">结构化导入错误。</param>
    /// <returns>能够安全使用低保留 HSSF 写出时返回 <see langword="true" />。</returns>
    internal static bool CanWrite(ExcelDataReaderFailureWorkbookCapture capture,
        IReadOnlyCollection<ExcelImportError> errors)
    {
        if (capture == null || errors == null)
            return false;

        foreach (var sheet in capture.Sheets)
        {
            if (sheet == null)
                return false;
            if (sheet.HasConditionalFormatting
                && !CanWriteConditionalFormattings(sheet))
                return false;
            if (!sheet.HasNativeDataValidation)
            {
                if (sheet.FreezePane != null && !CanWriteFreezePane(sheet.FreezePane))
                    return false;
                continue;
            }
            if (sheet.NativeDataValidations == null || sheet.NativeDataValidations.Count == 0
                || sheet.NativeDataValidations.Any(validation =>
                    !CanWriteNativeDataValidation(capture, sheet, validation)))
                return false;
            if (sheet.FreezePane != null && !CanWriteFreezePane(sheet.FreezePane))
                return false;
        }

        foreach (var error in errors)
        {
            if (error == null || error.RowIndex <= 0 || string.IsNullOrWhiteSpace(error.SheetName))
                return false;
            var sheet = capture.Sheets.FirstOrDefault(item =>
                string.Equals(item.Name, error.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet == null || !sheet.Rows.Any(row => row.RowIndex == error.RowIndex))
                return false;
        }
        return true;
    }

    /// <summary>
    /// 同步写出低保留 HSSF 失败工作簿。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">结构化导入错误。</param>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileExportCommitter">目标路径提交服务。</param>
    internal static void Write(ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None
            || errors == null || errors.Count == 0)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        var workbookAlreadyFiltered = HasNativeDataValidations(capture)
            || HasConditionalFormattings(capture);
        var requiredRowsBySheet = GetContextRowsBySheet(capture);
        using var workbook = CreateSparseWorkbook(capture, cancellationToken);
        NpoiFailureWorkbookWriter.Write(workbook, options, errors, EmptyResolvedSheetRequests,
            cancellationToken, new SystemFailureWorkbookFileSystem(), null,
            fileExportCommitter, requiredRowsBySheet,
            workbookAlreadyFiltered: workbookAlreadyFiltered);
    }

    /// <summary>
    /// 异步写出低保留 HSSF 失败工作簿。
    /// </summary>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">结构化导入错误。</param>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="fileExportCommitter">目标路径提交服务。</param>
    internal static async Task WriteAsync(ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken, IFileExportCommitter fileExportCommitter)
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None
            || errors == null || errors.Count == 0)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        var workbookAlreadyFiltered = HasNativeDataValidations(capture)
            || HasConditionalFormattings(capture);
        var requiredRowsBySheet = GetContextRowsBySheet(capture);
        using var workbook = CreateSparseWorkbook(capture, cancellationToken);
        await NpoiFailureWorkbookWriter.WriteAsync(workbook, options, errors,
            EmptyResolvedSheetRequests, cancellationToken, fileExportCommitter,
            requiredRowsBySheet,
            workbookAlreadyFiltered: workbookAlreadyFiltered)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// 获取命名列表来源所需的上下文行，供失败工作簿对象预算预检使用。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <returns>按工作表名称索引的源行集合。</returns>
    private static IReadOnlyDictionary<string, IReadOnlyList<int>> GetContextRowsBySheet(
        ExcelDataReaderFailureWorkbookCapture capture)
    {
        if (capture?.Sheets == null)
            return new Dictionary<string, IReadOnlyList<int>>(StringComparer.OrdinalIgnoreCase);
        return capture.Sheets
            .Where(sheet => sheet != null && sheet.ContextRows != null
                && sheet.ContextRows.Count > 0)
            .ToDictionary(sheet => sheet.Name,
                sheet => (IReadOnlyList<int>)sheet.ContextRows
                    .Where(row => row != null && row.RowIndex > 0)
                    .Select(row => row.RowIndex - 1)
                    .Distinct()
                    .ToArray(),
                StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 判断失败行快照是否已包含需要保留原始坐标的工作表规则。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <returns>包含原生规则时返回 <see langword="true" />。</returns>
    private static bool HasNativeDataValidations(
        ExcelDataReaderFailureWorkbookCapture capture)
        => capture?.Sheets.Any(sheet => sheet?.HasNativeDataValidation == true) == true;

    /// <summary>
    /// 判断失败行快照是否已包含需要保留的条件格式。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <returns>包含条件格式时返回 <see langword="true" />。</returns>
    private static bool HasConditionalFormattings(
        ExcelDataReaderFailureWorkbookCapture capture)
        => capture?.Sheets.Any(sheet => sheet?.HasConditionalFormatting == true) == true;

    /// <summary>
    /// 判断冻结窗格坐标是否能在 HSSF 工作簿中安全重建。
    /// </summary>
    /// <param name="freezePane">冻结窗格坐标。</param>
    /// <returns>坐标合法时返回 <see langword="true" />。</returns>
    private static bool CanWriteFreezePane(ExcelDataReaderFailureFreezePane freezePane)
        => freezePane != null
            && (freezePane.ColumnSplit > 0 || freezePane.RowSplit > 0)
            && freezePane.ColumnSplit >= 0 && freezePane.ColumnSplit <= 255
            && freezePane.RowSplit >= 0 && freezePane.RowSplit <= 65_535
            && freezePane.LeftColumn >= 0 && freezePane.LeftColumn <= 255
            && freezePane.TopRow >= 0 && freezePane.TopRow <= 65_535;

    /// <summary>
    /// 判断条件格式是否能在稀疏 HSSF 工作表中安全重建。
    /// </summary>
    /// <param name="sheet">失败行捕获工作表。</param>
    /// <returns>条件格式属于受支持的保留行范围时返回 <see langword="true" />。</returns>
    private static bool CanWriteConditionalFormattings(
        ExcelDataReaderFailureSheet sheet)
    {
        if (sheet == null || !sheet.HasConditionalFormatting
            || sheet.ConditionalFormattings == null
            || sheet.ConditionalFormattings.Count == 0)
            return false;
        var retainedRows = GetRetainedRows(sheet);
        foreach (var formatting in sheet.ConditionalFormattings)
        {
            if (formatting == null || formatting.Ranges == null
                || formatting.Ranges.Count == 0 || formatting.Rules == null
                || formatting.Rules.Count == 0 || formatting.Rules.Count > 3)
                return false;
            foreach (var range in formatting.Ranges)
            {
                if (range == null || range.FirstRow < 0 || range.LastRow < range.FirstRow
                    || range.LastRow > 65_535 || range.FirstColumn < 0
                    || range.LastColumn < range.FirstColumn || range.LastColumn > 255)
                    return false;
                for (var row = range.FirstRow; row <= range.LastRow; row++)
                {
                    if (!retainedRows.Contains(row))
                        return false;
                }
            }
            foreach (var rule in formatting.Rules)
            {
                if (!CanWriteConditionalFormattingRule(rule))
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// 判断条件格式规则是否能由 HSSF 高层 API 重建。
    /// </summary>
    /// <param name="rule">条件格式规则。</param>
    /// <returns>规则属于受支持的静态比较或公式子集时返回 <see langword="true" />。</returns>
    private static bool CanWriteConditionalFormattingRule(
        ExcelDataReaderFailureConditionalFormattingRule rule)
    {
        if (rule == null || (rule.ConditionType != 1 && rule.ConditionType != 2)
            || string.IsNullOrWhiteSpace(rule.Formula1) || rule.Formula1.Length > 255
            || rule.Formula1.IndexOf('!') >= 0 || rule.Formula1.IndexOf('[') >= 0
            || rule.Formula1.IndexOf(']') >= 0)
            return false;
        if (rule.ConditionType == 1
            && (rule.ComparisonOperator < 1 || rule.ComparisonOperator > 8
                || ((rule.ComparisonOperator == 1 || rule.ComparisonOperator == 2)
                    && string.IsNullOrWhiteSpace(rule.Formula2))))
            return false;
        var style = rule.FormattingStyle;
        if (style == null)
            return true;
        if (style.IsStrikeout || style.FontHeightInPoints > 409.5
            || (!string.IsNullOrWhiteSpace(style.FillPattern)
                && !string.Equals(style.FillPattern, "solid", StringComparison.Ordinal)))
            return false;
        return CanWriteConditionalBorder(style.BorderLeft)
            && CanWriteConditionalBorder(style.BorderRight)
            && CanWriteConditionalBorder(style.BorderTop)
            && CanWriteConditionalBorder(style.BorderBottom);
    }

    /// <summary>
    /// 判断条件格式边框是否在 HSSF 可表示范围内。
    /// </summary>
    /// <param name="side">边框状态。</param>
    /// <returns>边框为空或线型可表示时返回 <see langword="true" />。</returns>
    private static bool CanWriteConditionalBorder(ExcelDataReaderFailureBorderSide side)
        => side == null || ToBorderStyle(side.Style) != BorderStyle.None;

    /// <summary>
    /// 从捕获快照构造只包含表头和失败行的稀疏 HSSF 工作簿。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>待交给通用失败写出器的 HSSF 工作簿。</returns>
    private static HSSFWorkbook CreateSparseWorkbook(
        ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        var workbook = new HSSFWorkbook();
        var styles = new Dictionary<string, ICellStyle>(StringComparer.Ordinal);
        var dataFormat = workbook.CreateDataFormat();
        try
        {
            foreach (var sheet in capture?.Sheets ?? Array.Empty<ExcelDataReaderFailureSheet>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var destination = workbook.CreateSheet(sheet.Name);
                WriteRow(destination, sheet.HeaderRowIndex, sheet.Header, sheet.HeaderFormats,
                    sheet.HeaderBoldColumns, sheet.HeaderFontStyles, workbook, styles,
                    dataFormat);
                foreach (var failureRow in sheet.Rows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    WriteRow(destination, failureRow.RowIndex - 1, failureRow.Values,
                        failureRow.NumberFormats, failureRow.BoldColumns,
                        failureRow.FontStyles, workbook, styles, dataFormat);
                }
                var retainedRows = new HashSet<int>(sheet.Rows
                    .Where(row => row != null)
                    .Select(row => row.RowIndex - 1))
                {
                    sheet.HeaderRowIndex
                };
                foreach (var contextRow in sheet.ContextRows)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (contextRow == null || !retainedRows.Add(contextRow.RowIndex - 1))
                        continue;
                    WriteRow(destination, contextRow.RowIndex - 1, contextRow.Values,
                        contextRow.NumberFormats, contextRow.BoldColumns,
                        contextRow.FontStyles, workbook, styles, dataFormat);
                }
                ApplyColumnMetadata(destination, sheet.ColumnMetadata);
                ApplyRowMetadata(destination, sheet.HeaderRowIndex, sheet.Rows,
                    sheet.ContextRows,
                    sheet.RowMetadata);
                ApplyMergeRegions(destination, sheet.HeaderRowIndex, sheet.Rows,
                    sheet.ContextRows,
                    sheet.MergeRegions);
                ApplyFreezePane(destination, sheet.FreezePane);
                ApplyConditionalFormattings(destination, sheet, cancellationToken);
                var sheetIndex = workbook.GetSheetIndex(destination);
                if (string.Equals(sheet.Visibility, "veryHidden", StringComparison.OrdinalIgnoreCase))
                    workbook.SetSheetVisibility(sheetIndex, SheetVisibility.VeryHidden);
                else if (string.Equals(sheet.Visibility, "hidden", StringComparison.OrdinalIgnoreCase))
                    workbook.SetSheetVisibility(sheetIndex, SheetVisibility.Hidden);
            }
            ApplyDefinedNames(workbook, capture, cancellationToken);
            foreach (var sheet in capture?.Sheets ?? Array.Empty<ExcelDataReaderFailureSheet>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                ApplyNativeDataValidations(workbook.GetSheet(sheet.Name), sheet,
                    cancellationToken);
            }
            return workbook;
        }
        catch
        {
            workbook.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 写入能够在低保留工作簿中保持坐标有效的一维定义名称。
    /// </summary>
    /// <param name="workbook">待写入名称的 HSSF 工作簿。</param>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ApplyDefinedNames(HSSFWorkbook workbook,
        ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        if (workbook == null || capture?.Metadata?.DefinedNames == null)
            return;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definedName in capture.Metadata.DefinedNames)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (definedName == null || string.IsNullOrWhiteSpace(definedName.Name)
                || !TryBuildDefinedNameFormula(definedName.TargetSheetName,
                    definedName.Range, out var formula))
                continue;
            var targetSheet = workbook.GetSheet(definedName.TargetSheetName);
            if (targetSheet == null)
                continue;
            var localSheetIndex = -1;
            if (!string.IsNullOrWhiteSpace(definedName.LocalSheetName))
            {
                var localSheet = workbook.GetSheet(definedName.LocalSheetName);
                if (localSheet == null)
                    continue;
                localSheetIndex = workbook.GetSheetIndex(localSheet);
            }
            if (!seen.Add(definedName.Name + "\u001f" + localSheetIndex))
                continue;
            var name = workbook.CreateName();
            name.NameName = definedName.Name;
            name.SheetIndex = localSheetIndex;
            name.RefersToFormula = formula;
        }
    }

    /// <summary>
    /// 创建 HSSF 支持的一维绝对区域公式。
    /// </summary>
    /// <param name="sheetName">目标工作表名称。</param>
    /// <param name="range">零基源区域。</param>
    /// <param name="formula">创建的绝对区域公式。</param>
    /// <returns>区域合法且不超出 HSSF 边界时返回 <see langword="true" />。</returns>
    private static bool TryBuildDefinedNameFormula(string sheetName,
        ExcelXlsxFailureRange range, out string formula)
    {
        formula = null;
        if (string.IsNullOrWhiteSpace(sheetName)
            || (range.FirstRow != range.LastRow && range.FirstColumn != range.LastColumn)
            || range.FirstRow < 0 || range.LastRow < range.FirstRow
            || range.LastRow > 65_535 || range.FirstColumn < 0
            || range.LastColumn < range.FirstColumn || range.LastColumn > 255)
            return false;
        var first = AbsoluteCellReference(range.FirstRow, range.FirstColumn);
        var last = AbsoluteCellReference(range.LastRow, range.LastColumn);
        formula = QuoteSheet(sheetName) + "!" + first;
        if (!string.Equals(first, last, StringComparison.Ordinal))
            formula += ":" + last;
        return true;
    }

    /// <summary>
    /// 创建带绝对行列标记的 HSSF 单元格地址。
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
    private static string QuoteSheet(string sheetName)
        => "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'";

    /// <summary>
    /// 将捕获的值写入指定的 HSSF 行。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="rowIndex">零基行号。</param>
    /// <param name="values">连续单元格值。</param>
    /// <param name="numberFormats">单元格数字格式。</param>
    /// <param name="boldColumns">单元格加粗列。</param>
    /// <param name="fontStyles">单元格字体状态。</param>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="styles">按样式复用的缓存。</param>
    /// <param name="dataFormat">目标工作簿数字格式工厂。</param>
    private static void WriteRow(ISheet sheet, int rowIndex, IReadOnlyList<object> values,
        IReadOnlyDictionary<int, string> numberFormats,
        IReadOnlyDictionary<int, bool> boldColumns,
        IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> fontStyles,
        IWorkbook workbook, IDictionary<string, ICellStyle> styles, IDataFormat dataFormat)
    {
        if (sheet == null || rowIndex < 0 || values == null)
            return;
        var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
        for (var column = 0; column < values.Count; column++)
        {
            var value = values[column];
            if (value == null)
                continue;
            var cell = row.CreateCell(column);
            SetCellValue(cell, value);
            var format = numberFormats != null && numberFormats.TryGetValue(column, out var valueFormat)
                ? valueFormat : null;
            var isBold = boldColumns != null && boldColumns.ContainsKey(column);
            var fontStyle = fontStyles != null && fontStyles.TryGetValue(column, out var valueStyle)
                ? valueStyle : null;
            ApplyCellStyle(cell, format, isBold, fontStyle, workbook, styles, dataFormat);
        }
    }

    /// <summary>
    /// 为保留单元格应用安全的最小样式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="format">源工作簿数字格式。</param>
    /// <param name="isBold">是否应用粗体。</param>
    /// <param name="fontStyle">源工作簿字体状态。</param>
    /// <param name="workbook">目标工作簿。</param>
    /// <param name="styles">按格式复用的样式缓存。</param>
    /// <param name="dataFormat">目标工作簿数字格式工厂。</param>
    private static void ApplyCellStyle(ICell cell, string format, bool isBold,
        ExcelDataReaderFailureFontStyle fontStyle,
        IWorkbook workbook, IDictionary<string, ICellStyle> styles, IDataFormat dataFormat)
    {
        if (cell == null || workbook == null || styles == null || dataFormat == null)
            return;
        var hasFormat = !string.IsNullOrWhiteSpace(format)
            && !string.Equals(format, "General", StringComparison.OrdinalIgnoreCase)
            && format.Length <= 255;
        var hasFontStyle = isBold || fontStyle?.IsBold == true || fontStyle?.IsItalic == true
            || fontStyle?.IsStrikeout == true || fontStyle?.IsUnderline == true
            || !string.IsNullOrWhiteSpace(fontStyle?.FontName)
            || fontStyle?.FontHeightInPoints > 0
            || fontStyle?.FontColorIndex.HasValue == true
            || !string.IsNullOrWhiteSpace(fontStyle?.HorizontalAlignment)
             || !string.IsNullOrWhiteSpace(fontStyle?.VerticalAlignment)
             || fontStyle?.WrapText == true || fontStyle?.ShrinkToFit == true
             || fontStyle?.TextRotation.HasValue == true || fontStyle?.Indentation > 0
             || (string.Equals(fontStyle?.FillPattern, "solid", StringComparison.Ordinal)
                 && fontStyle.FillForegroundColorIndex.HasValue)
             || fontStyle?.BorderLeft != null || fontStyle?.BorderRight != null
             || fontStyle?.BorderTop != null || fontStyle?.BorderBottom != null
             || fontStyle?.IsLocked.HasValue == true || fontStyle?.IsHidden.HasValue == true;
        if (!hasFormat && !hasFontStyle)
            return;
        var key = (hasFormat ? format : string.Empty) + "|"
            + (isBold || fontStyle?.IsBold == true ? "1" : "0")
            + (fontStyle?.IsItalic == true ? "1" : "0")
            + (fontStyle?.IsStrikeout == true ? "1" : "0")
            + (fontStyle?.IsUnderline == true ? "1" : "0") + "|"
            + (fontStyle?.FontName ?? string.Empty) + "|"
            + (fontStyle?.FontHeightInPoints.ToString("R", CultureInfo.InvariantCulture) ?? "0") + "|"
            + (fontStyle?.FontColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
            + (fontStyle?.HorizontalAlignment ?? string.Empty) + "|"
            + (fontStyle?.VerticalAlignment ?? string.Empty) + "|"
            + (fontStyle?.WrapText == true ? "1" : "0")
            + (fontStyle?.ShrinkToFit == true ? "1" : "0") + "|"
             + (fontStyle?.TextRotation?.ToString(CultureInfo.InvariantCulture) ?? string.Empty) + "|"
             + (fontStyle?.Indentation.ToString(CultureInfo.InvariantCulture) ?? "0") + "|"
             + (fontStyle?.FillPattern ?? string.Empty) + "|"
             + (fontStyle?.FillForegroundColorIndex?.ToString(CultureInfo.InvariantCulture)
                 ?? string.Empty) + "|"
             + (fontStyle?.FillBackgroundColorIndex?.ToString(CultureInfo.InvariantCulture)
                 ?? string.Empty) + "|"
             + GetBorderKey(fontStyle?.BorderLeft) + "|"
             + GetBorderKey(fontStyle?.BorderRight) + "|"
             + GetBorderKey(fontStyle?.BorderTop) + "|"
             + GetBorderKey(fontStyle?.BorderBottom) + "|"
             + (fontStyle?.IsLocked?.ToString() ?? string.Empty) + "|"
             + (fontStyle?.IsHidden?.ToString() ?? string.Empty);
        if (!styles.TryGetValue(key, out var style))
        {
            style = workbook.CreateCellStyle();
            if (hasFormat)
                style.DataFormat = dataFormat.GetFormat(format);
            if (hasFontStyle)
            {
                var font = workbook.CreateFont();
                font.IsBold = isBold || fontStyle?.IsBold == true;
                font.IsItalic = fontStyle?.IsItalic == true;
                font.IsStrikeout = fontStyle?.IsStrikeout == true;
                if (!string.IsNullOrWhiteSpace(fontStyle?.FontName)
                    && fontStyle.FontName.Length <= 255)
                    font.FontName = fontStyle.FontName;
                if (fontStyle?.FontHeightInPoints > 0
                    && fontStyle.FontHeightInPoints <= 409.5)
                    font.FontHeightInPoints = (short)Math.Round(
                        fontStyle.FontHeightInPoints, MidpointRounding.AwayFromZero);
                if (fontStyle?.FontColorIndex.HasValue == true)
                    font.Color = fontStyle.FontColorIndex.Value;
                if (fontStyle?.IsUnderline == true)
                    font.Underline = FontUnderlineType.Single;
                style.SetFont(font);
            }
            if (!string.IsNullOrWhiteSpace(fontStyle?.HorizontalAlignment))
                style.Alignment = ToHorizontalAlignment(fontStyle.HorizontalAlignment);
            if (!string.IsNullOrWhiteSpace(fontStyle?.VerticalAlignment))
                style.VerticalAlignment = ToVerticalAlignment(fontStyle.VerticalAlignment);
            if (fontStyle?.WrapText == true)
                style.WrapText = true;
            if (fontStyle?.ShrinkToFit == true)
                style.ShrinkToFit = true;
            if (fontStyle?.TextRotation.HasValue == true)
                style.Rotation = fontStyle.TextRotation.Value;
             if (fontStyle?.Indentation > 0)
                 style.Indention = fontStyle.Indentation;
             if (string.Equals(fontStyle?.FillPattern, "solid", StringComparison.Ordinal)
                 && fontStyle.FillForegroundColorIndex.HasValue)
             {
                 style.FillPattern = FillPattern.SolidForeground;
                 style.FillForegroundColor = fontStyle.FillForegroundColorIndex.Value;
                 if (fontStyle.FillBackgroundColorIndex.HasValue)
                     style.FillBackgroundColor = fontStyle.FillBackgroundColorIndex.Value;
             }
             ApplyBorder(style, fontStyle?.BorderLeft, BorderSide.Left);
             ApplyBorder(style, fontStyle?.BorderRight, BorderSide.Right);
             ApplyBorder(style, fontStyle?.BorderTop, BorderSide.Top);
             ApplyBorder(style, fontStyle?.BorderBottom, BorderSide.Bottom);
             if (fontStyle?.IsLocked.HasValue == true)
                 style.IsLocked = fontStyle.IsLocked.Value;
             if (fontStyle?.IsHidden.HasValue == true)
                 style.IsHidden = fontStyle.IsHidden.Value;
             styles[key] = style;
        }
        cell.CellStyle = style;
    }

    /// <summary>
    /// 将 XLSX 水平对齐名称转换为 HSSF 对齐枚举。
    /// </summary>
    private static HorizontalAlignment ToHorizontalAlignment(string value)
        => value switch
        {
            "left" => HorizontalAlignment.Left,
            "center" => HorizontalAlignment.Center,
            "right" => HorizontalAlignment.Right,
            "fill" => HorizontalAlignment.Fill,
            "justify" => HorizontalAlignment.Justify,
            "centerContinuous" => HorizontalAlignment.CenterSelection,
            "distributed" => HorizontalAlignment.Distributed,
            _ => HorizontalAlignment.General
        };

    /// <summary>
    /// 将 XLSX 垂直对齐名称转换为 HSSF 对齐枚举。
    /// </summary>
    private static VerticalAlignment ToVerticalAlignment(string value)
        => value switch
        {
            "center" => VerticalAlignment.Center,
            "top" => VerticalAlignment.Top,
            "justify" => VerticalAlignment.Justify,
            "distributed" => VerticalAlignment.Distributed,
            _ => VerticalAlignment.Bottom
        };

    /// <summary>
    /// 获取边框状态缓存键。
    /// </summary>
    private static string GetBorderKey(ExcelDataReaderFailureBorderSide side)
        => side == null ? string.Empty : (side.Style ?? string.Empty) + ":"
            + (side.ColorIndex?.ToString(CultureInfo.InvariantCulture) ?? string.Empty);

    /// <summary>
    /// 将失败输出边框状态应用到 HSSF 样式。
    /// </summary>
    private static void ApplyBorder(ICellStyle style, ExcelDataReaderFailureBorderSide side,
        BorderSide borderSide)
    {
        if (style == null || side == null || string.IsNullOrWhiteSpace(side.Style))
            return;
        var borderStyle = ToBorderStyle(side.Style);
        switch (borderSide)
        {
            case BorderSide.Left:
                style.BorderLeft = borderStyle;
                if (side.ColorIndex.HasValue)
                    style.LeftBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Right:
                style.BorderRight = borderStyle;
                if (side.ColorIndex.HasValue)
                    style.RightBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Top:
                style.BorderTop = borderStyle;
                if (side.ColorIndex.HasValue)
                    style.TopBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Bottom:
                style.BorderBottom = borderStyle;
                if (side.ColorIndex.HasValue)
                    style.BottomBorderColor = side.ColorIndex.Value;
                break;
        }
    }

    /// <summary>
    /// 将 OOXML 边框线型转换为 HSSF 枚举。
    /// </summary>
    private static BorderStyle ToBorderStyle(string value)
        => value switch
        {
            "thin" => BorderStyle.Thin,
            "medium" => BorderStyle.Medium,
            "dashed" => BorderStyle.Dashed,
            "dotted" => BorderStyle.Dotted,
            "thick" => BorderStyle.Thick,
            "double" => BorderStyle.Double,
            "hair" => BorderStyle.Hair,
            "mediumDashed" => BorderStyle.MediumDashed,
            "dashDot" => BorderStyle.DashDot,
            "mediumDashDot" => BorderStyle.MediumDashDot,
            "dashDotDot" => BorderStyle.DashDotDot,
            "mediumDashDotDot" => BorderStyle.MediumDashDotDot,
            "slantDashDot" => BorderStyle.SlantedDashDot,
            _ => BorderStyle.None
        };

    /// <summary>
    /// 标识 HSSF 边框方向。
    /// </summary>
    private enum BorderSide
    {
        Left,
        Right,
        Top,
        Bottom
    }

    /// <summary>
    /// 应用捕获的列宽和隐藏状态。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="metadata">源工作表列元数据。</param>
    private static void ApplyColumnMetadata(ISheet sheet,
        IReadOnlyDictionary<int, ExcelDataReaderFailureColumnMetadata> metadata)
    {
        if (sheet == null || metadata == null)
            return;
        foreach (var pair in metadata)
        {
            if (pair.Key < 0 || pair.Key > 255)
                continue;
            var column = pair.Key;
            var item = pair.Value;
            if (item == null)
                continue;
            if (item.HasWidth)
            {
                var width = (int)Math.Round(item.Width * 256d,
                    MidpointRounding.AwayFromZero);
                width = Math.Max(0, Math.Min(255 * 256, width));
                sheet.SetColumnWidth(column, width);
            }
            sheet.SetColumnHidden(column, item.Hidden);
        }
    }

    /// <summary>
    /// 应用源工作表的冻结窗格。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="freezePane">源冻结窗格坐标。</param>
    private static void ApplyFreezePane(ISheet sheet,
        ExcelDataReaderFailureFreezePane freezePane)
    {
        if (sheet == null || !CanWriteFreezePane(freezePane))
            return;
        sheet.CreateFreezePane(freezePane.ColumnSplit, freezePane.RowSplit,
            freezePane.LeftColumn, freezePane.TopRow);
    }

    /// <summary>
    /// 将捕获的条件格式写入稀疏 HSSF 工作表。
    /// </summary>
    /// <param name="sheet">目标 HSSF 工作表。</param>
    /// <param name="capture">失败行捕获工作表。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ApplyConditionalFormattings(ISheet sheet,
        ExcelDataReaderFailureSheet capture, CancellationToken cancellationToken)
    {
        if (sheet == null || capture?.ConditionalFormattings == null
            || capture.ConditionalFormattings.Count == 0)
            return;
        var formatting = sheet.SheetConditionalFormatting;
        foreach (var captured in capture.ConditionalFormattings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!CanWriteConditionalFormattings(capture)
                || captured.Ranges == null || captured.Rules == null)
                continue;
            var ranges = captured.Ranges.Select(range => new CellRangeAddress(
                range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn)).ToArray();
            var rules = captured.Rules.Select(rule =>
                CreateConditionalFormattingRule(formatting, rule)).ToArray();
            if (ranges.Length == 0 || rules.Any(rule => rule == null))
                continue;
            foreach (var pair in captured.Rules.Zip(rules, (source, target) =>
                         (Source: source, Target: target)))
                ApplyConditionalFormattingStyle(pair.Target, pair.Source.FormattingStyle);
            formatting.AddConditionalFormatting(ranges, rules);
        }
    }

    /// <summary>
    /// 创建一条 HSSF 条件格式规则。
    /// </summary>
    /// <param name="formatting">目标工作表条件格式集合。</param>
    /// <param name="captured">捕获的规则。</param>
    /// <returns>创建的条件格式规则。</returns>
    private static IConditionalFormattingRule CreateConditionalFormattingRule(
        ISheetConditionalFormatting formatting,
        ExcelDataReaderFailureConditionalFormattingRule captured)
    {
        if (formatting == null || captured == null)
            return null;
        if (captured.ConditionType == 2)
            return formatting.CreateConditionalFormattingRule(captured.Formula1);
        if (!TryGetComparisonOperator(captured.ComparisonOperator, out var comparison))
            return null;
        return string.IsNullOrWhiteSpace(captured.Formula2)
            ? formatting.CreateConditionalFormattingRule(comparison, captured.Formula1)
            : formatting.CreateConditionalFormattingRule(comparison, captured.Formula1,
                captured.Formula2);
    }

    /// <summary>
    /// 将条件格式比较运算符转换为 NPOI 枚举。
    /// </summary>
    /// <param name="value">条件格式运算符数值。</param>
    /// <param name="comparison">转换后的枚举。</param>
    /// <returns>运算符在 HSSF 范围内时返回 <see langword="true" />。</returns>
    private static bool TryGetComparisonOperator(int value,
        out ComparisonOperator comparison)
    {
        comparison = value switch
        {
            1 => ComparisonOperator.Between,
            2 => ComparisonOperator.NotBetween,
            3 => ComparisonOperator.Equal,
            4 => ComparisonOperator.NotEqual,
            5 => ComparisonOperator.GreaterThan,
            6 => ComparisonOperator.LessThan,
            7 => ComparisonOperator.GreaterThanOrEqual,
            8 => ComparisonOperator.LessThanOrEqual,
            _ => ComparisonOperator.NoComparison
        };
        return value >= 1 && value <= 8;
    }

    /// <summary>
    /// 应用条件格式中的字体、填充和边框状态。
    /// </summary>
    /// <param name="rule">目标规则。</param>
    /// <param name="style">捕获的样式。</param>
    private static void ApplyConditionalFormattingStyle(IConditionalFormattingRule rule,
        ExcelDataReaderFailureFontStyle style)
    {
        if (rule == null || style == null)
            return;
        var font = rule.CreateFontFormatting();
        font.SetFontStyle(style.IsItalic, style.IsBold);
        font.UnderlineType = style.IsUnderline
            ? FontUnderlineType.Single : FontUnderlineType.None;
        if (style.FontHeightInPoints > 0)
            font.FontHeight = (int)Math.Round(style.FontHeightInPoints * 20d,
                MidpointRounding.AwayFromZero);
        if (style.FontColorIndex.HasValue)
            font.FontColorIndex = style.FontColorIndex.Value;

        if (string.Equals(style.FillPattern, "solid", StringComparison.Ordinal))
        {
            var pattern = rule.CreatePatternFormatting();
            pattern.FillPattern = FillPattern.SolidForeground;
            if (style.FillForegroundColorIndex.HasValue)
                pattern.FillForegroundColor = style.FillForegroundColorIndex.Value;
            if (style.FillBackgroundColorIndex.HasValue)
                pattern.FillBackgroundColor = style.FillBackgroundColorIndex.Value;
        }

        if (style.BorderLeft != null || style.BorderRight != null
            || style.BorderTop != null || style.BorderBottom != null)
        {
            var border = rule.CreateBorderFormatting();
            ApplyConditionalBorder(border, style.BorderLeft, BorderSide.Left);
            ApplyConditionalBorder(border, style.BorderRight, BorderSide.Right);
            ApplyConditionalBorder(border, style.BorderTop, BorderSide.Top);
            ApplyConditionalBorder(border, style.BorderBottom, BorderSide.Bottom);
        }
    }

    /// <summary>
    /// 应用条件格式的一侧边框。
    /// </summary>
    /// <param name="border">目标边框。</param>
    /// <param name="side">捕获的边框状态。</param>
    /// <param name="borderSide">边框方向。</param>
    private static void ApplyConditionalBorder(IBorderFormatting border,
        ExcelDataReaderFailureBorderSide side, BorderSide borderSide)
    {
        if (border == null || side == null)
            return;
        var style = ToBorderStyle(side.Style);
        switch (borderSide)
        {
            case BorderSide.Left:
                border.BorderLeft = style;
                if (side.ColorIndex.HasValue)
                    border.LeftBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Right:
                border.BorderRight = style;
                if (side.ColorIndex.HasValue)
                    border.RightBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Top:
                border.BorderTop = style;
                if (side.ColorIndex.HasValue)
                    border.TopBorderColor = side.ColorIndex.Value;
                break;
            case BorderSide.Bottom:
                border.BorderBottom = style;
                if (side.ColorIndex.HasValue)
                    border.BottomBorderColor = side.ColorIndex.Value;
                break;
        }
    }

    /// <summary>
    /// 应用表头和失败行的捕获行高及隐藏状态。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="headerRowIndex">表头的零基行号。</param>
    /// <param name="failureRows">失败行快照。</param>
    /// <param name="contextRows">命名区域上下文行快照。</param>
    /// <param name="metadata">源工作表行元数据。</param>
    private static void ApplyRowMetadata(ISheet sheet, int headerRowIndex,
        IReadOnlyList<ExcelDataReaderFailureRow> failureRows,
        IReadOnlyList<ExcelDataReaderFailureRow> contextRows,
        IReadOnlyDictionary<int, ExcelDataReaderFailureRowMetadata> metadata)
    {
        if (sheet == null || metadata == null)
            return;
        var rowIndexes = new HashSet<int> { headerRowIndex };
        if (failureRows != null)
        {
            foreach (var failureRow in failureRows)
            {
                if (failureRow != null)
                    rowIndexes.Add(failureRow.RowIndex - 1);
            }
        }
        if (contextRows != null)
        {
            foreach (var contextRow in contextRows)
            {
                if (contextRow != null)
                    rowIndexes.Add(contextRow.RowIndex - 1);
            }
        }
        foreach (var rowIndex in rowIndexes)
        {
            if (rowIndex < 0 || !metadata.TryGetValue(rowIndex, out var item) || item == null)
                continue;
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            if (item.HasHeight)
                row.Height = item.Height;
            row.ZeroHeight = item.Hidden;
        }
    }

    /// <summary>
    /// 应用完全位于保留行内的合并区域。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="headerRowIndex">表头的零基行号。</param>
    /// <param name="failureRows">失败行快照。</param>
    /// <param name="contextRows">命名区域上下文行快照。</param>
    /// <param name="mergeRegions">源工作表合并区域。</param>
    private static void ApplyMergeRegions(ISheet sheet, int headerRowIndex,
        IReadOnlyList<ExcelDataReaderFailureRow> failureRows,
        IReadOnlyList<ExcelDataReaderFailureRow> contextRows,
        IReadOnlyList<ExcelDataReaderFailureMergeRegion> mergeRegions)
    {
        if (sheet == null || mergeRegions == null || mergeRegions.Count == 0)
            return;
        var retainedRows = new HashSet<int> { headerRowIndex };
        if (failureRows != null)
        {
            foreach (var failureRow in failureRows)
            {
                if (failureRow != null)
                    retainedRows.Add(failureRow.RowIndex - 1);
            }
        }
        if (contextRows != null)
        {
            foreach (var contextRow in contextRows)
            {
                if (contextRow != null)
                    retainedRows.Add(contextRow.RowIndex - 1);
            }
        }
        foreach (var region in mergeRegions)
        {
            if (region == null || region.FirstRow < 0 || region.LastRow < region.FirstRow
                || region.FirstColumn < 0 || region.LastColumn < region.FirstColumn
                || region.LastRow > 65_535 || region.LastColumn > 255)
                continue;
            if ((long)region.LastRow - region.FirstRow + 1 > retainedRows.Count)
                continue;
            var allRowsRetained = true;
            for (var row = region.FirstRow; row <= region.LastRow; row++)
            {
                if (retainedRows.Contains(row))
                    continue;
                allRowsRetained = false;
                break;
            }
            if (!allRowsRetained || HasOverlappingMerge(sheet, region))
                continue;
            sheet.AddMergedRegion(new CellRangeAddress(region.FirstRow, region.LastRow,
                region.FirstColumn, region.LastColumn));
        }
    }

    /// <summary>
    /// 判断原生校验是否能直接写入保留的源行坐标。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="sheet">失败行捕获工作表。</param>
    /// <param name="validation">原生校验描述。</param>
    /// <returns>区域完整保留且约束可创建时返回 <see langword="true" />。</returns>
    private static bool CanWriteNativeDataValidation(
        ExcelDataReaderFailureWorkbookCapture capture,
        ExcelDataReaderFailureSheet sheet,
        ExcelDataReaderFailureNativeDataValidation validation)
    {
        if (sheet == null || validation == null || validation.FirstRow < 0
            || validation.LastRow < validation.FirstRow || validation.FirstColumn < 0
            || validation.LastColumn < validation.FirstColumn || validation.LastRow > 65_535
            || validation.LastColumn > 255 || !TryCreateValidationConstraint(validation, out _)
            || GetRetainedValidationRowSegments(sheet, validation).Count == 0)
            return false;
        if (validation.ValidationType == NPOI.SS.UserModel.ValidationType.LIST
            && (validation.ListValues == null || validation.ListValues.Count == 0)
            && !CanWriteNamedList(capture, validation))
            return false;
        return true;
    }

    /// <summary>
    /// 判断命名列表的来源名称和工作表是否可以安全重建。
    /// </summary>
    /// <param name="capture">失败行捕获快照。</param>
    /// <param name="validation">原生校验描述。</param>
    /// <returns>命名列表来源可以安全重建时返回 <see langword="true" />。</returns>
    private static bool CanWriteNamedList(
        ExcelDataReaderFailureWorkbookCapture capture,
        ExcelDataReaderFailureNativeDataValidation validation)
    {
        if (capture?.Metadata?.DefinedNames == null
            || string.IsNullOrWhiteSpace(validation?.ListSourceName))
            return false;
        var definedName = capture.Metadata.DefinedNames.FirstOrDefault(item =>
            item != null
            && string.Equals(item.Name, validation.ListSourceName,
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(item.LocalSheetName, validation.ListSourceLocalSheetName,
                StringComparison.OrdinalIgnoreCase));
        if (definedName == null
            || definedName.Range.FirstRow < 0 || definedName.Range.LastRow < definedName.Range.FirstRow
            || definedName.Range.FirstColumn < 0
            || definedName.Range.LastColumn < definedName.Range.FirstColumn
            || definedName.Range.LastRow > 65_535 || definedName.Range.LastColumn > 255)
            return false;
        var sourceSheet = capture.Sheets.FirstOrDefault(item =>
            item != null && string.Equals(item.Name, definedName.TargetSheetName,
                StringComparison.OrdinalIgnoreCase));
        if (sourceSheet == null)
            return false;
        return true;
    }

    /// <summary>
    /// 获取失败工作簿中仍保留的源行集合。
    /// </summary>
    /// <param name="sheet">失败行捕获工作表。</param>
    /// <returns>保留的零基行号集合。</returns>
    private static HashSet<int> GetRetainedRows(ExcelDataReaderFailureSheet sheet)
    {
        var rows = new HashSet<int> { sheet.HeaderRowIndex };
        foreach (var row in sheet.Rows)
        {
            if (row != null)
                rows.Add(row.RowIndex - 1);
        }
        foreach (var row in sheet.ContextRows)
        {
            if (row != null)
                rows.Add(row.RowIndex - 1);
        }
        return rows;
    }

    /// <summary>
    /// 将捕获的原生数值校验写入稀疏 HSSF 工作表。
    /// </summary>
    /// <param name="sheet">目标 HSSF 工作表。</param>
    /// <param name="capture">失败行捕获工作表。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ApplyNativeDataValidations(ISheet sheet,
        ExcelDataReaderFailureSheet capture, CancellationToken cancellationToken)
    {
        if (sheet == null || capture?.NativeDataValidations == null)
            return;
        var helper = sheet.GetDataValidationHelper();
        foreach (var validation in capture.NativeDataValidations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryCreateValidationConstraint(validation, out var constraint))
                continue;
            foreach (var segment in GetRetainedValidationRowSegments(capture, validation))
            {
                var region = new CellRangeAddressList(segment.FirstRow, segment.LastRow,
                    validation.FirstColumn, validation.LastColumn);
                var copied = helper.CreateValidation(constraint, region);
                copied.EmptyCellAllowed = validation.EmptyCellAllowed;
                copied.SuppressDropDownArrow = validation.SuppressDropdownArrow;
                copied.ShowPromptBox = validation.ShowPromptBox;
                if (validation.ShowPromptBox)
                    copied.CreatePromptBox(validation.PromptTitle, validation.PromptText);
                copied.ShowErrorBox = validation.ShowErrorBox;
                if (validation.ShowErrorBox)
                    copied.CreateErrorBox(validation.ErrorTitle, validation.ErrorText);
                copied.ErrorStyle = validation.ErrorStyle;
                sheet.AddValidationData(copied);
            }
        }
    }

    /// <summary>
    /// 获取原生校验区域与低保留输出行的连续交集。
    /// </summary>
    /// <param name="sheet">失败行捕获工作表。</param>
    /// <param name="validation">原生校验描述。</param>
    /// <returns>按源行坐标合并后的连续保留片段。</returns>
    private static IReadOnlyList<(int FirstRow, int LastRow)> GetRetainedValidationRowSegments(
        ExcelDataReaderFailureSheet sheet,
        ExcelDataReaderFailureNativeDataValidation validation)
    {
        if (sheet == null || validation == null)
            return Array.Empty<(int FirstRow, int LastRow)>();
        var rows = new HashSet<int>();
        if (sheet.HeaderRowIndex >= validation.FirstRow
            && sheet.HeaderRowIndex <= validation.LastRow)
            rows.Add(sheet.HeaderRowIndex);
        foreach (var row in sheet.Rows)
        {
            if (row == null)
                continue;
            var sourceRow = row.RowIndex - 1;
            if (sourceRow >= validation.FirstRow && sourceRow <= validation.LastRow)
                rows.Add(sourceRow);
        }
        foreach (var row in sheet.ContextRows)
        {
            if (row == null)
                continue;
            var sourceRow = row.RowIndex - 1;
            if (sourceRow >= validation.FirstRow && sourceRow <= validation.LastRow)
                rows.Add(sourceRow);
        }
        var ordered = rows.OrderBy(row => row).ToArray();
        if (ordered.Length == 0)
            return Array.Empty<(int FirstRow, int LastRow)>();
        var segments = new List<(int FirstRow, int LastRow)>();
        var first = ordered[0];
        var last = first;
        for (var index = 1; index < ordered.Length; index++)
        {
            if (ordered[index] == last + 1)
            {
                last = ordered[index];
                continue;
            }
            segments.Add((first, last));
            first = last = ordered[index];
        }
        segments.Add((first, last));
        return segments;
    }

    /// <summary>
    /// 创建可在稀疏 HSSF 工作表中重建的原生校验约束。
    /// </summary>
    /// <param name="validation">原生校验描述。</param>
    /// <param name="constraint">创建的约束。</param>
    /// <returns>约束可创建时返回 <see langword="true" />。</returns>
    private static bool TryCreateValidationConstraint(
        ExcelDataReaderFailureNativeDataValidation validation,
        out IDataValidationConstraint constraint)
    {
        constraint = null;
        if (validation == null)
            return false;
        if (validation.ValidationType == NPOI.SS.UserModel.ValidationType.LIST)
        {
            try
            {
                if (validation.ListValues != null && validation.ListValues.Count > 0)
                {
                    constraint = DVConstraint.CreateExplicitListConstraint(
                        validation.ListValues.ToArray());
                }
                else if (!string.IsNullOrWhiteSpace(validation.ListSourceName))
                {
                    constraint = DVConstraint.CreateFormulaListConstraint(
                        validation.ListSourceName);
                }
                return constraint != null;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
        if (validation.ValidationType != NPOI.SS.UserModel.ValidationType.INTEGER
            && validation.ValidationType != NPOI.SS.UserModel.ValidationType.DECIMAL
            && validation.ValidationType != NPOI.SS.UserModel.ValidationType.TEXT_LENGTH)
            return false;
        try
        {
            constraint = DVConstraint.CreateNumericConstraint(validation.ValidationType,
                validation.ComparisonOperator, validation.Formula1, validation.Formula2);
            return constraint != null;
        }
        catch (ArgumentException)
        {
            return false;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    /// <summary>
    /// 判断合并区域是否与已有区域相交。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="region">待检查的合并区域。</param>
    /// <returns>相交时返回 <see langword="true" />。</returns>
    private static bool HasOverlappingMerge(ISheet sheet,
        ExcelDataReaderFailureMergeRegion region)
    {
        for (var index = 0; index < sheet.NumMergedRegions; index++)
        {
            var existing = sheet.GetMergedRegion(index);
            if (existing.FirstRow <= region.LastRow && region.FirstRow <= existing.LastRow
                && existing.FirstColumn <= region.LastColumn
                && region.FirstColumn <= existing.LastColumn)
                return true;
        }
        return false;
    }

    /// <summary>
    /// 将对象值转换为 HSSF 支持的基础单元格类型。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="value">待写入值。</param>
    private static void SetCellValue(ICell cell, object value)
    {
        switch (value)
        {
            case string text:
                cell.SetCellValue(text);
                break;
            case bool boolean:
                cell.SetCellValue(boolean);
                break;
            case DateTime dateTime:
                cell.SetCellValue(dateTime);
                break;
            case DateTimeOffset dateTimeOffset:
                cell.SetCellValue(dateTimeOffset.DateTime);
                break;
            case byte number:
                cell.SetCellValue(number);
                break;
            case sbyte number:
                cell.SetCellValue(number);
                break;
            case short number:
                cell.SetCellValue(number);
                break;
            case ushort number:
                cell.SetCellValue(number);
                break;
            case int number:
                cell.SetCellValue(number);
                break;
            case uint number:
                cell.SetCellValue(number);
                break;
            case long number:
                cell.SetCellValue(number);
                break;
            case ulong number:
                cell.SetCellValue(number);
                break;
            case float number:
                cell.SetCellValue(number);
                break;
            case double number:
                cell.SetCellValue(number);
                break;
            case decimal number:
                cell.SetCellValue((double)number);
                break;
            default:
                cell.SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture)
                    ?? string.Empty);
                break;
        }
    }
}
