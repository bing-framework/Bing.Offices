using System.ComponentModel;

namespace Bing.Offices.Imports;

/// <summary>单表便利入口的导入结果，保留 Workbook 管线的错误、源行和错误截断信息。</summary>
/// <typeparam name="T">工作表中的实体类型。</typeparam>
public sealed class ExcelSingleSheetImportResult<T>
{
    /// <summary>使用既有导入管线的结果创建单表视图，不复制实体或错误集合。</summary>
    /// <param name="items">成功导入的实体集合。</param>
    /// <param name="errors">完整结构化错误，包括缺少 Sheet 的 Workbook 级错误。</param>
    /// <param name="sourceRows">与成功实体逐项对应的零基源行索引。</param>
    /// <param name="errorsTruncated">错误收集是否达到上限。</param>
    /// <param name="maxErrors">配置的错误收集上限。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public ExcelSingleSheetImportResult(IReadOnlyList<T> items, IReadOnlyList<ExcelImportError> errors,
        IReadOnlyList<int> sourceRows, bool errorsTruncated, int? maxErrors)
    {
        Items = items ?? throw new ArgumentNullException(nameof(items));
        Errors = errors ?? throw new ArgumentNullException(nameof(errors));
        SourceRows = sourceRows ?? throw new ArgumentNullException(nameof(sourceRows));
        ErrorsTruncated = errorsTruncated;
        MaxErrors = maxErrors;
    }

    /// <summary>获取成功导入的实体。</summary>
    public IReadOnlyList<T> Items { get; }

    /// <summary>获取原始 Workbook 管线提供的全部错误。</summary>
    public IReadOnlyList<ExcelImportError> Errors { get; }

    /// <summary>获取与 Items 逐项对应的零基源行；用于用户界面时加 1。</summary>
    public IReadOnlyList<int> SourceRows { get; }

    /// <summary>获取错误收集是否达到上限，不表示所有资源截断。</summary>
    public bool ErrorsTruncated { get; }

    /// <summary>获取配置的错误收集上限。</summary>
    public int? MaxErrors { get; }

    /// <summary>获取是否没有错误且结果完整。</summary>
    public bool IsSuccess => Errors.Count == 0 && IsComplete;

    /// <summary>获取是否没有结构拒绝、资源限制或错误截断；普通行错误和显式配置的范围终点不影响此标志。</summary>
    public bool IsComplete => StopReason == ImportStopReason.None;

    /// <summary>获取不完整结果的原因摘要。</summary>
    public ImportStopReason StopReason => ImportResultState.GetReason(Errors, ErrorsTruncated);
}
