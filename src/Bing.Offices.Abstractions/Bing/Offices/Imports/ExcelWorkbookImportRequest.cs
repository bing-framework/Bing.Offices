using System.Linq.Expressions;
using System.ComponentModel;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// Workbook 导入请求。
/// </summary>
/// <typeparam name="TWorkbook">承载各工作表实体的工作簿模型类型。</typeparam>
public sealed class ExcelWorkbookImportRequest<TWorkbook> where TWorkbook : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookImportRequest{TWorkbook}" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按配置顺序排列的工作表请求。</param>
    /// <param name="relations">工作簿父子关系请求。</param>
    /// <param name="sheetNameComparison">工作表名称比较策略。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="failureOptions">失败工作簿输出选项。</param>
    /// <param name="validationMode">导入校验模式。</param>
    /// <param name="unsupportedFeaturePolicy">不支持功能的处理策略。</param>
    internal ExcelWorkbookImportRequest(IReadOnlyList<ExcelSheetImportRequest> sheets,
        IReadOnlyList<ExcelRelationRequest> relations, ExcelNameComparison sheetNameComparison,
        ExcelResourceLimits resourceLimits, ExcelImportFailureOptions failureOptions,
        ExcelImportValidationMode validationMode, ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy)
    {
        Sheets = sheets;
        Relations = relations;
        SheetNameComparison = sheetNameComparison;
        ResourceLimits = resourceLimits;
        FailureOptions = failureOptions;
        ValidationMode = validationMode;
        UnsupportedFeaturePolicy = unsupportedFeaturePolicy;
    }

    /// <summary>
    /// 获取Sheet 配置数量。
    /// </summary>
    public int SheetCount => Sheets.Count;

    /// <summary>
    /// 获取不可变 Sheet 导入描述。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelSheetImportRequest> Sheets { get; }
    /// <summary>
    /// 获取父子关系描述。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public IReadOnlyList<ExcelRelationRequest> Relations { get; }
    /// <summary>
    /// 获取Sheet 名称比较策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelNameComparison SheetNameComparison { get; }
    /// <summary>
    /// 获取输入资源限制。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelResourceLimits ResourceLimits { get; }
    /// <summary>
    /// 获取失败工作簿选项。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportFailureOptions FailureOptions { get; }
    /// <summary>
    /// 获取校验模式。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelImportValidationMode ValidationMode { get; }
    /// <summary>
    /// 获取不支持特性策略。
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; }
}
