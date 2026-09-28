using System.Globalization;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;

namespace Bing.Offices.Dates;

/// <summary>
/// 无 offset 文本转换为 DateTimeOffset 时使用的策略。
/// </summary>
public enum ExcelDateOffsetPolicy
{
    /// <summary>
    /// 必须由输入文本显式提供 offset。
    /// </summary>
    RequireExplicitOffset,
    /// <summary>
    /// 使用配置的固定 offset。
    /// </summary>
    UseFixedOffset
}
