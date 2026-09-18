using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 数据注解测试
/// </summary>
public class ExportTestDataAnnotations
{
    /// <summary>
    /// 获取或设置第一个名称测试值。
    /// </summary>
    public string Name1 { get; set; }

    /// <summary>
    /// 获取或设置第二个名称测试值。
    /// </summary>
    public string Name2 { get; set; }

    /// <summary>
    /// 获取或设置使用标准日期格式的时间测试值。
    /// </summary>
    [DisplayFormat(DataFormatString = "yyyy-MM-dd")]
    public DateTime Time1 { get; set; }

    /// <summary>
    /// 获取或设置使用自定义日期格式的时间测试值。
    /// </summary>
    [DataFormat("yyyy-MM-dd")]
    public DateTime Time2 { get; set; }

    /// <summary>
    /// 获取或设置应忽略的测试值。
    /// </summary>
    [ExcelIgnore]
    public string Ignore { get; set; }

    /// <summary>
    /// 获取或设置值映射枚举测试值。
    /// </summary>
    [ValueMapping("A Test", 0)]
    [ValueMapping("B Test", 1)]
    public MyEnum MyEnum { get; set; }

    /// <summary>
    /// 获取或设置可空布尔值映射测试值。
    /// </summary>
    [ValueMapping("是", true)]
    [ValueMapping("否", false)]
    public bool? Bool { get; set; }

    /// <summary>
    /// 获取或设置布尔值映射测试值。
    /// </summary>
    [ValueMapping("是", true)]
    [ValueMapping("否", false)]
    public bool Bool1 { get; set; }

    /// <summary>
    /// 获取或设置可空布尔值测试值。
    /// </summary>
    public bool? Bool2 { get; set; }

    /// <summary>
    /// 获取或设置可空整数测试值。
    /// </summary>
    public int? Number { get; set; }

    /// <summary>
    /// 获取或设置带小数位配置的数值测试值。
    /// </summary>
    [DecimalScale(1)]
    public decimal Scale { get; set; }

    /// <summary>
    /// 获取或设置使用货币格式的数值测试值。
    /// </summary>
    [DisplayFormat(DataFormatString = "C")]
    public decimal Scale1 { get; set; }
}

/// <summary>
/// 表示数据注解测试使用的值映射枚举。
/// </summary>
public enum MyEnum
{
    /// <summary>
    /// 表示值映射测试中的 A 值。
    /// </summary>
    A,
    /// <summary>
    /// 表示值映射测试中的 B 值。
    /// </summary>
    B,
    /// <summary>
    /// 表示带有“C Test”描述的值映射测试值。
    /// </summary>
    [Description("C Test")]
    C,
    /// <summary>
    /// 表示未配置额外描述的值映射测试值。
    /// </summary>
    D
}
