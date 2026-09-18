using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs;

/// <summary>
/// 表示 Issue8 场景使用的 Excel 测试模型。
/// </summary>
public class Issue8
{
    /// <summary>
    /// 获取或设置编码。
    /// </summary>
    [ColumnName("编码")]
    [ExcelRequired(ErrorMsg = "编码为必填项")]
    public string Code { get; set; }

    /// <summary>
    /// 获取或设置可选时间。
    /// </summary>
    [ColumnName("非必填时间")]
    public dynamic CreateTime { get; set; }
}
