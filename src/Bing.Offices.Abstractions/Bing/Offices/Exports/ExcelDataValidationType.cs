namespace Bing.Offices.Exports;

/// <summary>
/// 工作表原生数据校验类型。
/// </summary>
/// <remarks>用于生成工作表规则，不执行导入校验。</remarks>
public enum ExcelDataValidationType
{
    /// <summary>
    /// 整数。
    /// </summary>
    Integer,
    /// <summary>
    /// 小数。
    /// </summary>
    Decimal,
    /// <summary>
    /// 日期。
    /// </summary>
    Date,
    /// <summary>
    /// 文本长度。
    /// </summary>
    TextLength,
    /// <summary>
    /// 显式列表。
    /// </summary>
    List
}
