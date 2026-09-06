using System;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导入车牌
/// </summary>
public class ImportCar
{
    [ColumnName("车牌号")]
    [ExcelRequired]
    [ExcelRegex(@"^[京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤青藏川宁琼使领A-Z]{1}[A-Z]{1}[A-Z0-9]{4}[A-Z0-9挂学警港澳]{1}$")]
    [ExcelUnique]
    public string CarCode { get; set; }

    [ColumnName("手机号")]
    [ExcelRegex(@"0?(13|14|15|17|18|19)[0-9]{9}")]
    public string Mobile { get; set; }

    [ColumnName("身份证")]
    [ExcelRegex(@"^(^\d{15}$|^\d{18}$|^\d{17}(\d|X|x))$")]
    public string IdCard { get; set; }

    [ColumnName("姓名")]
    [ExcelMaxLength(10)]
    public string Name { get; set; }

    [ColumnName("性别")]
    [ExcelRegex(@"^男$|^女$|^先生$|^女士$|^Male|^Female$")]
    public Gender Gender { get; set; }

    [ColumnName("注册日期")]
    [ExcelDate]
    public DateTime RegisterDate { get; set; }

    [ColumnName("年龄")]
    [ExcelRange(0,150)]
    public int Age { get; set; }
}
