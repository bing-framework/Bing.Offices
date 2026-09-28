namespace Bing.Offices.Imports;

/// <summary>
/// 跟踪 Excel 导入过程中的行数和图片资源配额。
/// </summary>
internal sealed class ExcelImportRuntime
{
    /// <summary>
    /// 记录当前工作簿已尝试处理的数据行数量。
    /// </summary>
    private int _rowCount;
    /// <summary>
    /// 指示行数上限错误是否已加入结果，避免重复报告。
    /// </summary>
    private bool _rowLimitReported;

    /// <summary>
    /// 初始化一个 <see cref="ExcelImportRuntime" /> 类型的实例。
    /// </summary>
    /// <param name="limits">导入请求配置的资源限制。</param>
    internal ExcelImportRuntime(ExcelResourceLimits limits)
    {
        MaxRows = limits?.MaxRows;
        ImageResources = new ExcelImageResourceTracker(limits);
    }

    /// <summary>
    /// 获取允许处理的最大数据行数。
    /// </summary>
    internal int? MaxRows { get; }

    /// <summary>
    /// 获取是否已达到数据行数上限。
    /// </summary>
    internal bool RowLimitReached => MaxRows.HasValue && _rowCount >= MaxRows.Value;

    /// <summary>
    /// 获取是否已发现超出工作簿行数预算的数据行。
    /// </summary>
    internal bool RowLimitExceeded => _rowLimitReported;

    /// <summary>
    /// 获取跟踪工作簿图片数量和字节数的资源限制器。
    /// </summary>
    internal ExcelImageResourceTracker ImageResources { get; }

    /// <summary>
    /// 尝试为一行数据消耗全局行数配额。
    /// </summary>
    /// <returns>成功消耗配额时为 true；已达到上限时为 false。</returns>
    internal bool TryConsumeRow()
    {
        if (MaxRows.HasValue && _rowCount >= MaxRows.Value)
            return false;
        _rowCount++;
        return true;
    }

    /// <summary>
    /// 确保行数上限错误在单个工作簿中仅报告一次。
    /// </summary>
    /// <returns>本次调用首次标记上限错误时为 true。</returns>
    internal bool TryMarkRowLimitReported()
    {
        if (_rowLimitReported)
            return false;
        _rowLimitReported = true;
        return true;
    }
}
