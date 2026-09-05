using System;
using Bing.Offices.Dates;

namespace Bing.Offices.Attributes;

/// <summary>
/// v2 日期校验特性。
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public sealed class ExcelDateAttribute : FilterAttributeBase
{
    /// <summary>
    /// 使用精确日期格式初始化校验特性。
    /// </summary>
    /// <param name="format">日期格式。</param>
    public ExcelDateAttribute(string format) => Format = format;

    /// <summary>
    /// 使用默认 ISO 日期格式初始化校验特性。
    /// </summary>
    public ExcelDateAttribute()
    {
    }

    /// <summary>
    /// 获取或设置精确输入日期格式；为空时使用 yyyy-MM-dd。
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// 获取或设置解析日期时使用的区域性名称。
    /// </summary>
    public string CultureName { get; set; }

    /// <summary>获取或设置无 offset 文本的 DateTimeOffset 解析策略。</summary>
    public ExcelDateOffsetPolicy OffsetPolicy { get; set; } = ExcelDateOffsetPolicy.RequireExplicitOffset;

    /// <summary>获取或设置固定 offset 分钟数；仅 UseFixedOffset 策略使用。</summary>
    public int? OffsetMinutes { get; set; }

    /// <inheritdoc />
    public override string ErrorMsg { get; set; } = "非日期数据";
}
