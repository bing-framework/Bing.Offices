namespace Bing.Offices.Csv;

/// <summary>
/// CSV 导入错误分类。
/// </summary>
public enum CsvImportErrorCode
{
    /// <summary>
    /// 输入格式无效。
    /// </summary>
    InvalidInput,
    /// <summary>
    /// 表头或列结构无效。
    /// </summary>
    InvalidHeader,
    /// <summary>
    /// 值转换失败。
    /// </summary>
    ValueConversion,
    /// <summary>
    /// 业务校验失败。
    /// </summary>
    Validation,
    /// <summary>
    /// 超过资源限制。
    /// </summary>
    ResourceLimit
}
