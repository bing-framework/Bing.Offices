using System;
using System.ComponentModel;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导入测试样例
/// </summary>
public class ImportSample
{
    /// <summary>
    /// 获取或设置测试值。
    /// </summary>
    public string TestValue { get; set; }

    /// <summary>
    /// 获取或设置电子邮件验证值。
    /// </summary>
    public string Email { get; set; }

    /// <summary>
    /// 获取或设置网址验证值。
    /// </summary>
    public string Url { get; set; }

    /// <summary>
    /// 获取或设置最大长度测试值。
    /// </summary>
    public string MaxLengthValue { get; set; }

    /// <summary>
    /// 获取或设置 decimal 值。
    /// </summary>
    public decimal DecimalValue { get; set; }

    /// <summary>
    /// 获取或设置可空 decimal 值。
    /// </summary>
    public decimal? NullableDecimalValue { get; set; }

    /// <summary>
    /// 获取或设置 float 值。
    /// </summary>
    public float FloatValue { get; set; }

    /// <summary>
    /// 获取或设置可空 float 值。
    /// </summary>
    public float? NullableFloatValue { get; set; }

    /// <summary>
    /// 获取或设置 double 值。
    /// </summary>
    public double DoubleValue { get; set; }

    /// <summary>
    /// 获取或设置可空 double 值。
    /// </summary>
    public double? NullableDoubleValue { get; set; }

    /// <summary>
    /// 获取或设置 bool 值。
    /// </summary>
    public bool BoolValue { get; set; }

    /// <summary>
    /// 获取或设置可空 bool 值。
    /// </summary>
    public bool? NullableBoolValue { get; set; }

    /// <summary>
    /// 获取或设置 DateTime 值。
    /// </summary>
    public DateTime DateValue { get; set; }

    /// <summary>
    /// 获取或设置可空 DateTime 值。
    /// </summary>
    public DateTime? NullableDateValue { get; set; }

    /// <summary>
    /// 获取或设置不可空枚举值。
    /// </summary>
    public Gender LogLevel { get; set; }

    /// <summary>
    /// 获取或设置可空枚举值。
    /// </summary>
    public Gender? NullableLogLevel { get; set; }

    /// <summary>
    /// 获取或设置 int 值。
    /// </summary>
    [Description("IntValue")]
    public int IntValue { get; set; }

    /// <summary>
    /// 获取或设置可空 int 值。
    /// </summary>
    public int? NullableIntValue { get; set; }

    /// <summary>
    /// 获取或设置 short 值。
    /// </summary>
    public short ShortValue { get; set; }

    /// <summary>
    /// 获取或设置可空 short 值。
    /// </summary>
    public short? NullableShortValue { get; set; }

    /// <summary>
    /// 获取或设置 long 值。
    /// </summary>
    public long LongValue { get; set; }

    /// <summary>
    /// 获取或设置可空 long 值。
    /// </summary>
    public long? NullableLongValue { get; set; }

    /// <summary>
    /// 获取或设置应忽略的值。
    /// </summary>
    public string IgnoreValue { get; set; }
}
