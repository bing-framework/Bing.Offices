using Bing.Offices.Exceptions;
using Bing.Offices.Formula;

namespace Bing.Offices.ClosedXml.Formula;

/// <summary>
/// ClosedXML 公式读取适配器。
/// </summary>
/// <remarks>
/// 只读取公式文本和文件中已有的缓存值。ClosedXML 计算引擎并不等同于完整 Excel
/// 计算引擎，因此本适配器不接受重新计算请求。
/// </remarks>
public sealed class ClosedXmlExcelFormulaProcessor : IExcelFormulaProcessor
{
    /// <inheritdoc />
    public ExcelFormulaResult Process(Stream source, ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("输入流不可读取。", nameof(source));
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        ValidateRequest(request);
        if (format != ExcelFormat.Xlsx)
            throw Unsupported("ClosedXML 公式读取只支持 XLSX。", BingOfficesStage.Preflight);
        if (request.CalculationMode != ExcelFormulaCalculationMode.DoNotCalculate)
            throw Unsupported("ClosedXML Provider 只读取公式文本和缓存值，不承诺重新计算。",
                BingOfficesStage.Preflight);

        using var buffered = new MemoryStream();
        source.CopyTo(buffered);
        buffered.Position = 0;
        using var workbook = new ClosedXML.Excel.XLWorkbook(buffered);
        var cells = new List<ExcelFormulaCellResult>();
        foreach (var worksheet in workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(request.SheetName)
                && !string.Equals(worksheet.Name, request.SheetName, StringComparison.OrdinalIgnoreCase))
                continue;
            foreach (var row in worksheet.RowsUsed())
            {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (var cell in row.CellsUsed())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var formula = cell.FormulaA1;
                    if (string.IsNullOrWhiteSpace(formula))
                        continue;
                    cells.Add(new ExcelFormulaCellResult
                    {
                        SheetName = worksheet.Name,
                        RowIndex = cell.Address.RowNumber - 1,
                        ColumnIndex = cell.Address.ColumnNumber - 1,
                        Formula = request.ReadMode == ExcelFormulaReadMode.CachedValue ? null : formula,
                        CachedValue = request.ReadMode == ExcelFormulaReadMode.FormulaText
                            ? null : ToCachedValue(cell.CachedValue)
                    });
                }
            }
        }

        return new ExcelFormulaResult { Cells = cells, CalculationCompleted = true };
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
    /// 创建不支持当前功能的异常。
    /// </summary>
    /// <param name="message">错误消息。</param>
    /// <param name="stage">发生错误的处理阶段。</param>
    /// <returns>包含当前处理阶段的功能不支持异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message,
        BingOfficesStage stage) => new(message, provider: "ClosedXML",
        operation: BingOfficesOperation.Import, stage: stage);

    /// <summary>
    /// 转换公式缓存值。
    /// </summary>
    /// <param name="value">待转换的值。</param>
    /// <returns>转换后的缓存值；空白缓存值返回 null。</returns>
    private static object ToCachedValue(ClosedXML.Excel.XLCellValue value)
    {
        if (value.IsBlank) return null;
        if (value.IsBoolean) return value.GetBoolean();
        if (value.IsNumber) return value.GetNumber();
        if (value.IsDateTime) return value.GetDateTime();
        if (value.IsTimeSpan) return value.GetTimeSpan();
        if (value.IsError) return value.GetError().ToString();
        return value.GetText();
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
}
