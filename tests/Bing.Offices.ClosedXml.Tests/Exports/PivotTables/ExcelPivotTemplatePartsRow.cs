namespace Bing.Offices.ClosedXml.Tests.Exports.PivotTables;

/// <summary>
/// 表示透视模板部件测试使用的完整来源行。
/// </summary>
internal sealed class ExcelPivotTemplatePartsRow
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplatePartsRow"/> 类型的实例。
    /// </summary>
    public ExcelPivotTemplatePartsRow()
    {
    }

    /// <summary>
    /// 获取或设置区域名称。
    /// </summary>
    public string Region { get; set; }

    /// <summary>
    /// 获取或设置业务金额。
    /// </summary>
    public double Amount { get; set; }
}
