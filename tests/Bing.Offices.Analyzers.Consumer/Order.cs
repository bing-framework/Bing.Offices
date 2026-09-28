using Bing.Offices.Entities;

namespace Bing.Offices.Analyzers.Consumer;

/// <summary>
/// 供独立 NuGet 消费验证使用的采购单。
/// </summary>
public sealed class Order
{
    /// <summary>
    /// 获取或设置采购单编号。
    /// </summary>
#if BOE_NEGATIVE
    [ExcelEntityCell("Data", "A0")]
#else
    [ExcelEntityCell("Data", "A1")]
#endif
    public string Code { get; set; } = string.Empty;
}

#if BOE_AREA_NEGATIVE
/// <summary>
/// 验证打包分析器拒绝合并区域重叠的布局。
/// </summary>
public static class InvalidAreaLayout
{
    /// <summary>
    /// 构建重叠合并区域的负例布局。
    /// </summary>
    public static ExcelEntityLayout<Order> Build() => ExcelEntity.Layout<Order>(builder => builder
        .Cell("Data", "F1", item => item.Code)
        .Merge("Data", "A1:C3")
        .Merge("Data", "C3:E5"));
}
#endif
