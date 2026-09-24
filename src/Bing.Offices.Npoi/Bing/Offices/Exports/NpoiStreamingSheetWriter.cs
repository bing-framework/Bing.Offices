using System.Collections;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Providers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Exports;

/// <summary>
/// 为大型、无高级布局的 XLSX 列表执行流式行写入。
/// </summary>
/// <remarks>
/// SXSSF 不能在行刷新后再执行合并、自动列宽或图表等回溯操作；这些请求由调用方排除后才进入此写入器。
/// </remarks>
internal sealed class NpoiStreamingSheetWriter
{
    /// <summary>
    /// 将一个已编译映射计划按顺序写入流式工作表。
    /// </summary>
    /// <typeparam name="T">工作表数据项类型。</typeparam>
    /// <param name="workbook">流式 NPOI 工作簿。</param>
    /// <param name="request">当前工作表导出请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="mapping">当前工作表的映射计划。</param>
    /// <param name="columns">按导出顺序排列的列执行计划。</param>
    /// <param name="originRow">区域起始行；大型纯列表请求固定为零。</param>
    /// <param name="originColumn">区域起始列；大型纯列表请求固定为零。</param>
    internal void Write<T>(IWorkbook workbook, ExcelSheetExportRequest request,
        CancellationToken cancellationToken, IExcelMappingPlan mapping,
        IReadOnlyList<ExcelColumnPlan> columns, int originRow, int originColumn)
        where T : class, new()
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        if (request == null)
            throw new ArgumentNullException(nameof(request));

        var sheet = workbook.GetSheet(request.Name) ?? workbook.CreateSheet(request.Name);
        if (request.Hidden)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(sheet), SheetVisibility.Hidden);

        var mappingHeaderStyle = NpoiExportSheetWriter.ResolveStyle(mapping.Style?.HeaderStyleKey, true);
        var mappingBodyStyle = NpoiExportSheetWriter.ResolveStyle(mapping.Style?.BodyStyleKey, false);
        var dynamicKeys = request.FailOnUnknownDynamicValues
            ? new HashSet<string>(columns.Where(column => column.IsDynamic).Select(column => column.Key),
                StringComparer.Ordinal)
            : null;
        var header = sheet.CreateRow(originRow + request.HeaderRowIndex);
        var headerAttribute = typeof(T).GetCustomAttributes(typeof(HeaderAttribute), false)
            .Cast<HeaderAttribute>().SingleOrDefault();
        for (var index = 0; index < columns.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = header.CreateCell(originColumn + index);
            var style = columns[index].HeaderStyle ?? request.HeaderStyle ?? mappingHeaderStyle ?? request.SheetStyle;
            if (style != null)
                cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, style);
            if (headerAttribute != null)
                cell.CellStyle = NpoiStyleCache.ApplyHeaderAttribute(workbook, cell.CellStyle, headerAttribute);
            cell.SetCellValue(columns[index].Title);
        }

        var rowIndex = originRow + request.DataRowStartIndex;
        foreach (var item in request.Data.Cast<T>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = sheet.CreateRow(rowIndex);
            IDictionary<string, object> dynamicValues;
            try
            {
                dynamicValues = request.DynamicGetter?.Invoke(item);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                var dynamicProperty = columns.FirstOrDefault(column => column.IsDynamic)?.Property.Name;
                throw new BingOfficesExportException("Excel 动态值读取器执行失败。", exception, "NPOI",
                    BingOfficesStage.Validate, sheet.SheetName, rowIndex + 1, null, dynamicProperty,
                    BingOfficesErrorCode.UserExtensionFailed);
            }
            NpoiExportSheetWriter.ValidateUnknownDynamicValues(request, dynamicValues, dynamicKeys);
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var physicalColumnIndex = originColumn + columnIndex;
                var cell = row.CreateCell(physicalColumnIndex);
                var style = columns[columnIndex].BodyStyle ?? request.BodyStyle ?? mappingBodyStyle
                    ?? request.SheetStyle;
                NpoiExportSheetWriter.WriteCell(cell, item, columns[columnIndex], dynamicValues,
                    sheet.SheetName, rowIndex + 1, physicalColumnIndex + 1, request.Culture);
                if (style != null)
                    cell.CellStyle = NpoiStyleCache.Compose(workbook, cell.CellStyle, style);
            }
            rowIndex++;
        }
        if (request.RowHeight != null)
        {
            request.RowHeight.Validate();
            if (request.RowHeight.HeaderHeight.HasValue)
                (sheet.GetRow(originRow + request.HeaderRowIndex) ??
                    sheet.CreateRow(originRow + request.HeaderRowIndex)).HeightInPoints =
                    (float)request.RowHeight.HeaderHeight.Value;
            if (request.RowHeight.BodyHeight.HasValue && rowIndex > originRow + request.DataRowStartIndex)
            {
                for (var bodyRow = originRow + request.DataRowStartIndex; bodyRow < rowIndex; bodyRow++)
                    (sheet.GetRow(bodyRow) ?? sheet.CreateRow(bodyRow)).HeightInPoints =
                        (float)request.RowHeight.BodyHeight.Value;
            }
        }
    }
}
