using System.Globalization;
using Bing.Offices.Exceptions;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// Failure Workbook 的汇总工作表和错误批注职责。
/// </summary>
internal static class NpoiFailureWorkbookAnnotationWriter
{
    /// <summary>
    /// 将导入错误写入失败工作簿汇总工作表。
    /// </summary>
    /// <param name="workbook">待写入汇总工作表的工作簿。</param>
    /// <param name="errors">待写入的导入错误集合。</param>
    /// <param name="cancellationToken">写入过程中检查的取消令牌。</param>
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

    /// <summary>
    /// 将导入错误作为批注写回原始工作表单元格。
    /// </summary>
    /// <param name="workbook">包含原始工作表的工作簿。</param>
    /// <param name="errors">待标注的导入错误集合。</param>
    /// <param name="conflictPolicy">已有批注时的处理策略。</param>
    internal static void AnnotateErrors(IWorkbook workbook, IReadOnlyCollection<ExcelImportError> errors,
        ExcelImportCommentConflictPolicy conflictPolicy)
    {
        foreach (var error in errors.Where(error => error.RowIndex > 0 && error.ColumnIndex > 0))
        {
            var sheet = workbook.GetSheet(error.SheetName);
            if (sheet == null)
                continue;
            var row = sheet.GetRow(error.RowIndex - 1) ?? sheet.CreateRow(error.RowIndex - 1);
            var cell = row.GetCell(error.ColumnIndex - 1) ?? row.CreateCell(error.ColumnIndex - 1);
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

    /// <summary>
    /// 将错误原始值转换为汇总表可写入的文本。
    /// </summary>
    /// <param name="value">错误记录中的原始值。</param>
    /// <returns>用于汇总表的文本；二进制值显示其字节数。</returns>
    private static string FormatRawValue(object value) => value is byte[] bytes
        ? $"<binary:{bytes.Length}>"
        : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
}
