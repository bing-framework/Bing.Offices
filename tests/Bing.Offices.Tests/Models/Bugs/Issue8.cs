using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs;

public class Issue8
{
    /// <summary>
    /// 编码
    /// </summary>
    [ColumnName("编码")]
    [Required(ErrorMsg = "编码为必填项")]
    public string Code { get; set; }

    /// <summary>
    /// 非必填时间
    /// </summary>
    [ColumnName("非必填时间")]
    public dynamic CreateTime { get; set; }
}