using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Metadata;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;

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
    {
        if (options == null || options.Mode == ExcelImportFailureWorkbookMode.None || errors.Count == 0)
            return;
        cancellationToken.ThrowIfCancellationRequested();
        IWorkbook outputWorkbook = workbook;
        IWorkbook independentWorkbook = null;
        try
        {
            if (options.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly)
                outputWorkbook = independentWorkbook = CreateErrorRowsWorkbook(workbook, errors, resolvedSheetRequests,
                    cancellationToken);
            else
                AnnotateErrors(outputWorkbook, errors);

            WriteFailureSummary(outputWorkbook, errors, cancellationToken);
            using var output = new MemoryStream();
            outputWorkbook.Write(output, false);
            var bytes = output.ToArray();
            if (options.MaxBytes.HasValue && bytes.LongLength > options.MaxBytes.Value)
                throw new InvalidOperationException($"失败工作簿超过最大字节数: {options.MaxBytes.Value}");
            WriteBytes(options.Destination, bytes, cancellationToken);
        }
        finally
        {
            independentWorkbook?.Close();
        }
    }

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
                CopyRow(sourceSheet.GetRow(pair.Key), destinationSheet.CreateRow(pair.Value), destination,
                    styleCache);
            }
            AddFailureColumns(destinationSheet, group.Value, rowMap,
                sourceSheet.GetRow(headerRowIndex)?.LastCellNum ?? 0);
            CopyMergedRegions(sourceSheet, destinationSheet, rowMap);
            CopyDataValidations(sourceSheet, destinationSheet, rowMap, cancellationToken);
            CopyPictures(sourceSheet, destinationSheet, rowMap, cancellationToken);
        }
        return destination;
    }

    private static void CopyRow(IRow sourceRow, IRow destinationRow, IWorkbook destination,
        IDictionary<short, ICellStyle> styleCache)
    {
        if (sourceRow == null)
            return;
        destinationRow.Height = sourceRow.Height;
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
                    destinationCell.SetCellValue(sourceCell.StringCellValue);
                    break;
            }
            if (sourceCell.CellStyle != null)
            {
                if (!styleCache.TryGetValue(sourceCell.CellStyle.Index, out var style))
                {
                    style = destination.CreateCellStyle();
                    style.CloneStyleFrom(sourceCell.CellStyle);
                    styleCache[sourceCell.CellStyle.Index] = style;
                }
                destinationCell.CellStyle = style;
            }
        }
    }

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
                destination.AddValidationData(helper.CreateValidation(validation.ValidationConstraint,
                    targetRegion));
            }
        }
    }

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

    private static string FormatRawValue(object value)
    {
        if (value is byte[] bytes)
            return $"<binary:{bytes.Length}>";
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
    }

    private static void AnnotateErrors(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors)
    {
        foreach (var error in errors.Where(error => error.RowIndex > 0 && error.ColumnIndex > 0))
        {
            var sheet = workbook.GetSheet(error.SheetName);
            var cell = sheet?.GetRow(error.RowIndex - 1)?.GetCell(error.ColumnIndex - 1);
            if (cell == null)
                continue;
            var existing = cell.CellComment;
            var text = error.Message ?? string.Empty;
            if (existing != null)
                text = existing.String.String + Environment.NewLine + text;
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

    private static void WriteBytes(Stream destination, byte[] bytes, CancellationToken cancellationToken)
    {
        if (destination == null || !destination.CanWrite)
            throw new ArgumentException("失败工作簿目标流不可写入。", nameof(destination));
        var offset = 0;
        while (offset < bytes.Length)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var count = Math.Min(81920, bytes.Length - offset);
            destination.Write(bytes, offset, count);
            offset += count;
        }
    }
}
