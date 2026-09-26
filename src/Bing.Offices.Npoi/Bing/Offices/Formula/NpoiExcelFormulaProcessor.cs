using Bing.Offices.Exceptions;
using Bing.Offices.Formula;
using Bing.Offices.Providers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Formula;

/// <summary>
/// NPOI 公式读取适配器。
/// </summary>
/// <remarks>
/// 当前只读取公式文本和工作簿中已有的缓存值。NPOI 的公式计算器不作为 Excel
/// 完整兼容引擎暴露；重新计算请求在预检阶段显式拒绝，避免把不完整计算当成成功结果。
/// </remarks>
public sealed class NpoiExcelFormulaProcessor : IExcelFormulaProcessor
{
    /// <inheritdoc />
    public ExcelFormulaResult Process(Stream source, ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("输入流不可读取。", nameof(source));
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        ValidateFormat(format);
        ValidateRequest(request);
        if (request.CalculationMode != ExcelFormulaCalculationMode.DoNotCalculate)
            throw Unsupported("NPOI Provider 只读取公式文本和缓存值，不承诺重新计算。",
                BingOfficesStage.Preflight);

        using var buffered = new MemoryStream();
        source.CopyTo(buffered);
        buffered.Position = 0;
        using var workbook = WorkbookFactory.Create(buffered);
        var cells = new List<ExcelFormulaCellResult>();
        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sheet = workbook.GetSheetAt(sheetIndex);
            if (!string.IsNullOrWhiteSpace(request.SheetName)
                && !string.Equals(sheet.SheetName, request.SheetName, StringComparison.OrdinalIgnoreCase))
                continue;
            for (var rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var row = sheet.GetRow(rowIndex);
                if (row == null || row.FirstCellNum < 0)
                    continue;
                for (var columnIndex = row.FirstCellNum; columnIndex < row.LastCellNum; columnIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = row.GetCell(columnIndex);
                    if (cell == null)
                        continue;
                    if (cell.CellType != CellType.Formula)
                        continue;
                    cells.Add(new ExcelFormulaCellResult
                    {
                        SheetName = sheet.SheetName,
                        RowIndex = rowIndex,
                        ColumnIndex = cell.ColumnIndex,
                        Formula = request.ReadMode == ExcelFormulaReadMode.CachedValue
                            ? null : cell.CellFormula,
                        CachedValue = request.ReadMode == ExcelFormulaReadMode.FormulaText
                            ? null : GetCachedValue(cell),
                        ErrorCode = cell.CachedFormulaResultType == CellType.Error
                            ? FormulaError.ForInt(cell.ErrorCellValue).String : null
                    });
                }
            }
        }

        return new ExcelFormulaResult
        {
            Cells = cells,
            CalculationCompleted = true
        };
    }

    /// <inheritdoc />
    public async Task<ExcelFormulaResult> ProcessAsync(Stream source, ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("输入流不可读取。", nameof(source));
        await using var buffered = new MemoryStream();
        await source.CopyToAsync(buffered, 81920, cancellationToken).ConfigureAwait(false);
        buffered.Position = 0;
        return Process(buffered, format, request, cancellationToken);
    }

    /// <summary>
    /// 读取公式缓存值。
    /// </summary>
    /// <param name="cell">待处理的单元格。</param>
    /// <returns>缓存值或错误文本；无可读取的缓存类型时返回 null。</returns>
    private static object GetCachedValue(ICell cell)
    {
        switch (cell.CachedFormulaResultType)
        {
            case CellType.Numeric:
                return cell.NumericCellValue;
            case CellType.Boolean:
                return cell.BooleanCellValue;
            case CellType.String:
                return cell.StringCellValue;
            case CellType.Error:
                return FormulaError.ForInt(cell.ErrorCellValue).String;
            default:
                return null;
        }
    }

    /// <summary>
    /// 校验工作簿格式是否受支持。
    /// </summary>
    /// <param name="format">工作簿格式。</param>
    private static void ValidateFormat(ExcelFormat format)
    {
        if (format != ExcelFormat.Xls && format != ExcelFormat.Xlsx)
            throw Unsupported("NPOI 公式读取只支持 XLS/XLSX。", BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 校验公式读取与计算模式。
    /// </summary>
    /// <param name="request">当前操作请求。</param>
    private static void ValidateRequest(ExcelFormulaRequest request)
    {
        if (!Enum.IsDefined(typeof(ExcelFormulaReadMode), request.ReadMode))
            throw new ArgumentOutOfRangeException(nameof(request.ReadMode));
        if (!Enum.IsDefined(typeof(ExcelFormulaCalculationMode), request.CalculationMode))
            throw new ArgumentOutOfRangeException(nameof(request.CalculationMode));
    }

    /// <summary>
    /// 创建不支持当前功能的异常。
    /// </summary>
    /// <param name="message">错误消息。</param>
    /// <param name="stage">发生错误的处理阶段。</param>
    /// <returns>包含当前处理阶段的功能不支持异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message,
        BingOfficesStage stage) => new(message, provider: "NPOI",
        operation: BingOfficesOperation.Import, stage: stage);
}
