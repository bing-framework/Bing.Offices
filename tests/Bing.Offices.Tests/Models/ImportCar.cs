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
    [ExcelRegex(RegexConst.ChinaCarCode)]
    [ExcelUnique]
    public string CarCode { get; set; }

    [ColumnName("手机号")]
    [ExcelRegex(RegexConst.ChinaMobile)]
    public string Mobile { get; set; }

    [ColumnName("身份证")]
    [ExcelRegex(RegexConst.IdCard)]
    public string IdCard { get; set; }

    [ColumnName("姓名")]
    [ExcelMaxLength(10)]
    public string Name { get; set; }

    [ColumnName("性别")]
    [ExcelRegex(RegexConst.Gender)]
    public Gender Gender { get; set; }

    [ColumnName("注册日期")]
    [ExcelDate]
    public DateTime RegisterDate { get; set; }

    [ColumnName("年龄")]
    [ExcelRange(0,150)]
    public int Age { get; set; }
}
