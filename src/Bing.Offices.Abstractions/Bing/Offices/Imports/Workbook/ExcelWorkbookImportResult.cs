using System.ComponentModel;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入结果。
/// </summary>
/// <typeparam name="TWorkbook">承载各工作表实体的工作簿模型类型。</typeparam>
public sealed class ExcelWorkbookImportResult<TWorkbook> where TWorkbook : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookImportResult{TWorkbook}" /> 类型的实例。
    /// </summary>
    /// <param name="workbook">导入得到的根 Workbook 模型。</param>
    /// <param name="sheets">各工作表导入结果。</param>
    /// <param name="errors">关系和结构化导入错误。</param>
    /// <param name="errorsTruncated">是否因错误数量上限而停止收集。</param>
    /// <param name="maxErrors">生效的最大错误数；未限制时为 <see langword="null" />。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelWorkbookImportResult(TWorkbook workbook, IReadOnlyList<ExcelSheetImportResult> sheets,
        IReadOnlyList<ExcelImportError> errors, bool errorsTruncated, int? maxErrors)
    {
        Workbook = workbook;
        Sheets = sheets;
        Errors = errors;
        ErrorsTruncated = errorsTruncated;
        MaxErrors = maxErrors;
    }

    /// <summary>
    /// 获取根 Workbook 模型。
    /// </summary>
    public TWorkbook Workbook { get; }

    /// <summary>
    /// 获取各 Sheet 结果。
    /// </summary>
    public IReadOnlyList<ExcelSheetImportResult> Sheets { get; }

    /// <summary>
    /// 获取关系和结构化导入错误。
    /// </summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>
    /// 获取是否因错误数量上限而停止收集后续错误。
    /// </summary>
    public bool ErrorsTruncated { get; }

    /// <summary>
    /// 获取生效的最大错误数。
    /// </summary>
    public int? MaxErrors { get; }

    /// <summary>
    /// 获取是否没有错误。
    /// </summary>
    public bool IsSuccess => Errors.Count == 0;
}
