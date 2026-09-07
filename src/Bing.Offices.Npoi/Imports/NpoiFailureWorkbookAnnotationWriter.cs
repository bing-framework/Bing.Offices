using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>Failure Workbook 的汇总工作表和错误批注职责。</summary>
internal static class NpoiFailureWorkbookAnnotationWriter
{
    internal static void WriteSummary(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
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

    internal static void AnnotateErrors(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
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
                throw new BingOfficesConfigurationException($"单元格已有失败批注目标: {cell.Address}",
                    stage: BingOfficesStage.Plan);
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

    private static string FormatRawValue(object value) => value is byte[] bytes
        ? $"<binary:{bytes.Length}>"
        : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
}
