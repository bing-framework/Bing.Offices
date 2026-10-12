namespace Bing.Offices.Entities;

/// <summary>
/// 实体布局分批导入的解析模式。
/// </summary>
public enum ExcelEntityBatchParsingMode
{
    /// <summary>
    /// 使用 Provider 的完整工作簿解析路径。
    /// </summary>
    Buffered = 0,

    /// <summary>
    /// 使用受能力边界约束的前向解析路径。
    /// </summary>
    /// <remarks>
    /// Provider 只在明确支持的布局和文件格式下启用该模式；不支持时返回结构化配置错误，
    /// 不会静默退回完整工作簿解析。
    /// </remarks>
    Forward = 1
}
