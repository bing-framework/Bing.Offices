namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 表示模板透视合同测试中的业务数据行。
/// </summary>
internal sealed class TemplatePivotBoundaryRow
{
    /// <summary>
    /// 初始化一个 <see cref="TemplatePivotBoundaryRow" /> 类型的实例。
    /// </summary>
    public TemplatePivotBoundaryRow()
    {
    }

    /// <summary>
    /// 获取或设置行标签。
    /// </summary>
    public string Region { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置度量数值。
    /// </summary>
    public double Amount { get; set; }
}
