using System;

namespace Bing.Offices.Entities;

/// <summary>
/// 指定实体分批导入完成后保留的结果范围。
/// </summary>
public enum ExcelEntityBatchResultMode
{
    /// <summary>
    /// 保留完整实体结果和来源行索引。
    /// </summary>
    Complete = 0,

    /// <summary>
    /// 仅保留批次统计、结构化错误和工作表诊断。
    /// </summary>
    /// <remarks>
    /// 最终结果中的实体为空聚合对象；批次回调仍接收读取期间填充的实体。
    /// </remarks>
    SummaryOnly = 1
}
