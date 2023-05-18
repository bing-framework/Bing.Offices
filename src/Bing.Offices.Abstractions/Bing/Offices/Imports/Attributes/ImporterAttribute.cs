namespace Bing.Offices.Imports.Attributes;

/// <summary>
/// 导入 特性
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class ImporterAttribute : Attribute
{
    /// <summary>
    /// 表头位置
    /// </summary>
    public int HeaderRowIndex { get; set; } = 1;

    /// <summary>
    /// 最大允许导入的行数
    /// </summary>
    public int MaxCount = 0;

    /// <summary>
    /// 导入结果筛选器
    /// </summary>
    public Type ImportResultFilter { get; set; }

    /// <summary>
    /// 导入列头筛选器
    /// </summary>
    public Type ImportHeaderFilter { get; set; }

    /// <summary>
    /// 是否禁用所有筛选器
    /// </summary>
    public bool IsDisableAllFilter { get; set; }

    /// <summary>
    /// 是否忽略列的大小写
    /// </summary>
    public bool IsIgnoreColumnCase { get; set; }
}
