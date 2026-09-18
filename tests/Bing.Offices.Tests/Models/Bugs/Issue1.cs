using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models.Bugs;

/// <summary>
/// 表示 Issue1 场景使用的 Excel 测试模型。
/// </summary>
public class Issue1
{
    /// <summary>
    /// 获取或设置待测试的十进制值。
    /// </summary>
    [ColumnName("值")]
    public decimal DecimalValue { get; set; }
}
