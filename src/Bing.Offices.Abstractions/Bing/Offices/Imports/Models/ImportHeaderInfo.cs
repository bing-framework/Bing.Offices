namespace Bing.Offices.Imports.Models;

/// <summary>
/// 导入列头设置
/// </summary>
public class ImportHeaderInfo
{
    /// <summary>
    /// 是否必填
    /// </summary>
    public bool IsRequired { get; set; }

    /// <summary>
    /// 列名称
    /// </summary>
    public string PropertyName { get; set; }
}
