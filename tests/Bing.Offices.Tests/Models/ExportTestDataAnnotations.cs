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
    public string Name1 { get; set; }

    public string Name2 { get; set; }

    [DisplayFormat(DataFormatString = "yyyy-MM-dd")]
    public DateTime Time1 { get; set; }

    [DataFormat("yyyy-MM-dd")]
    public DateTime Time2 { get; set; }

    [ExcelIgnore]
    public string Ignore { get; set; }

    [ValueMapping("A Test", 0)]
    [ValueMapping("B Test", 1)]
    public MyEnum MyEnum { get; set; }

    [ValueMapping("是", true)]
    [ValueMapping("否", false)]
    public bool? Bool { get; set; }

    [ValueMapping("是", true)]
    [ValueMapping("否", false)]
    public bool Bool1 { get; set; }

    public bool? Bool2 { get; set; }

    public int? Number { get; set; }

    [DecimalScale(1)]
    public decimal Scale { get; set; }

    [DisplayFormat(DataFormatString = "C")]
    public decimal Scale1 { get; set; }
}

public enum MyEnum
{
    A,
    B,
    [Description("C Test")]
    C,
    D
}