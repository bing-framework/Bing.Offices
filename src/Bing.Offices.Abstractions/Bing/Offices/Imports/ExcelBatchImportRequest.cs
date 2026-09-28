using System.Globalization;
using Bing.Offices.Configurations;

namespace Bing.Offices.Imports;

/// <summary>
/// 单个分批导入请求。
/// </summary>
/// <typeparam name="TItem">工作表实体类型。</typeparam>
public sealed class ExcelBatchImportRequest<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelBatchImportRequest{TItem}"/> 类型的实例。
    /// </summary>
    /// <param name="sheetName">要读取的工作表名称。</param>
    public ExcelBatchImportRequest(string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(sheetName));
        SheetName = sheetName;
    }

    /// <summary>
    /// 获取工作表名称。
    /// </summary>
    public string SheetName { get; }
    /// <summary>
    /// 获取或初始化表头行的零基索引。
    /// </summary>
    public int HeaderRowIndex { get; init; }
    /// <summary>
    /// 获取或初始化数据起始行的零基索引。
    /// </summary>
    public int DataRowStartIndex { get; init; } = 1;
    /// <summary>
    /// 获取或初始化单批最大实体数量。
    /// </summary>
    public int BatchSize { get; init; } = 1000;
    /// <summary>
    /// 获取或初始化是否要求映射列全部存在。
    /// </summary>
    public bool RequireExpectedHeaders { get; init; } = true;
    /// <summary>
    /// 获取或初始化校验失败处理模式。
    /// </summary>
    public ExcelValidationFailureMode ValidationFailureMode { get; init; } = ExcelValidationFailureMode.StopOnFirstFailure;
    /// <summary>
    /// 获取或初始化映射方向的校验模式。
    /// </summary>
    public ExcelImportValidationMode ValidationMode { get; init; } = ExcelImportValidationMode.ConfiguredRules;
    /// <summary>
    /// 获取或初始化转换区域性。
    /// </summary>
    public CultureInfo Culture { get; init; } = CultureInfo.InvariantCulture;
    /// <summary>
    /// 获取或初始化请求级映射配置。
    /// </summary>
    public ExcelMappingConfiguration MappingConfiguration { get; init; }
    /// <summary>
    /// 获取或初始化规范化映射文档。
    /// </summary>
    public ExcelMappingDocument MappingDocument { get; init; }
    /// <summary>
    /// 获取或初始化导入资源限制。
    /// </summary>
    public ExcelResourceLimits ResourceLimits { get; init; } = new ExcelResourceLimits();
    /// <summary>
    /// 获取或初始化表头比较策略。
    /// </summary>
    public ExcelNameComparison HeaderComparison { get; init; } = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>
    /// 获取或初始化表头空白处理策略。
    /// </summary>
    public ExcelWhitespacePolicy HeaderWhitespace { get; init; } = ExcelWhitespacePolicy.Trim;
    /// <summary>
    /// 获取或初始化正文空白处理策略。
    /// </summary>
    public ExcelWhitespacePolicy BodyWhitespace { get; init; } = ExcelWhitespacePolicy.Trim;

    /// <summary>
    /// 验证请求参数。
    /// </summary>
    public void Validate()
    {
        if (HeaderRowIndex < 0 || DataRowStartIndex <= HeaderRowIndex)
            throw new ArgumentOutOfRangeException(nameof(DataRowStartIndex));
        if (BatchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(BatchSize));
        if (!Enum.IsDefined(typeof(ExcelValidationFailureMode), ValidationFailureMode))
            throw new ArgumentOutOfRangeException(nameof(ValidationFailureMode));
        if (!Enum.IsDefined(typeof(ExcelImportValidationMode), ValidationMode))
            throw new ArgumentOutOfRangeException(nameof(ValidationMode));
        if (Culture == null)
            throw new ArgumentNullException(nameof(Culture));
        ResourceLimits?.Validate();
    }
}
