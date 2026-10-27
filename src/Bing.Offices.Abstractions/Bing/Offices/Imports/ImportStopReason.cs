namespace Bing.Offices.Imports;

/// <summary>导入结果中结构拒绝、资源限制或错误收集截断的统一摘要。</summary>
/// <remarks>多个原因同时存在时，资源限制优先，其次错误上限、表头结构和未指定原因；详见结果 Errors。</remarks>
public enum ImportStopReason
{
    /// <summary>没有结构或限制导致的不完整结果；仍可能存在普通行校验错误。显式配置的范围终点也属于此状态。</summary>
    None,
    /// <summary>至少一个表头或 Sheet 结构被拒绝，无法按配置完成该部分导入。</summary>
    InvalidHeader,
    /// <summary>至少一项资源限制导致部分输入无法处理。</summary>
    ResourceLimit,
    /// <summary>错误收集达到上限，后续处理或错误收集可能未完成。</summary>
    ErrorLimit,
    /// <summary>结果标记为截断，但没有足够的结构化诊断确定原因。</summary>
    Unspecified,
    /// <summary>流式消费者主动停止，后续数据尚未读取。</summary>
    ConsumerStopped
}

/// <summary>从既有结构化诊断归纳结果状态，不改变导入处理策略。</summary>
internal static class ImportResultState
{
    /// <summary>归纳 Excel 结构化错误和错误截断标志。</summary>
    internal static ImportStopReason GetReason(IReadOnlyList<ExcelImportError> errors, bool errorsTruncated)
    {
        if (errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit))
            return ImportStopReason.ResourceLimit;
        if (errorsTruncated)
            return ImportStopReason.ErrorLimit;
        return errors.Any(error => error.Code == ExcelImportErrorCode.InvalidHeader)
            ? ImportStopReason.InvalidHeader : ImportStopReason.None;
    }

    /// <summary>归纳 CSV 结构化错误，保留旧截断标志与错误上限的不同含义。</summary>
    internal static ImportStopReason GetReason(IReadOnlyList<Csv.CsvImportError> errors, bool isTruncated,
        int? maxErrors)
    {
        if (errors.Any(error => error.Code == Csv.CsvImportErrorCode.ResourceLimit))
            return ImportStopReason.ResourceLimit;
        if (isTruncated && maxErrors.HasValue && errors.Count >= maxErrors.Value)
            return ImportStopReason.ErrorLimit;
        if (errors.Any(error => error.Code == Csv.CsvImportErrorCode.InvalidHeader))
            return ImportStopReason.InvalidHeader;
        return isTruncated ? ImportStopReason.Unspecified : ImportStopReason.None;
    }
}
