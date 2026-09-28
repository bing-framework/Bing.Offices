using System.ComponentModel;

namespace Bing.Offices.Formula;

/// <summary>
/// 单个公式结果。
/// </summary>
public sealed class ExcelFormulaCellResult
{
    /// <summary>
    /// 获取或初始化Sheet 名称。
    /// </summary>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化零基行索引。
    /// </summary>
    public int RowIndex { get; init; }
    /// <summary>
    /// 获取或初始化零基列索引。
    /// </summary>
    public int ColumnIndex { get; init; }
    /// <summary>
    /// 获取或初始化公式文本；未请求时为空。
    /// </summary>
    public string Formula { get; init; }
    /// <summary>
    /// 获取或初始化缓存值；未请求或不存在时为空。
    /// </summary>
    public object CachedValue { get; init; }
    /// <summary>
    /// 获取或初始化计算后值。
    /// </summary>
    public object CalculatedValue { get; init; }
    /// <summary>
    /// 获取或初始化结构化错误代码。
    /// </summary>
    public string ErrorCode { get; init; }
}
