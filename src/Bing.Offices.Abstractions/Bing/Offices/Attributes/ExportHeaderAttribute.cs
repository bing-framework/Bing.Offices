namespace Bing.Offices.Attributes;

/// <summary>
/// 导出头 特性
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExportHeaderAttribute : Attribute
{
    /// <summary>
    /// 显示名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 备注
    /// </summary>
    public string Comment { get; set; }

    /// <summary>
    /// 字体大小
    /// </summary>
    public float FontSize { get; set; } = 12;

    /// <summary>
    /// 是否加粗
    /// </summary>
    public bool IsBold { get; set; } = true;

    /// <summary>
    /// 格式化字符串
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// 是否自适应
    /// </summary>
    public bool IsAutoFit { get; set; } = false;

    /// <summary>
    /// 自动居中
    /// </summary>
    public bool AutoCenterColumn { get; set; }

    /// <summary>
    /// 是否忽略
    /// </summary>
    public bool IsIgnore { get; set; }

    /// <summary>
    /// 宽度，<see cref="IsAutoFit"/> 为 <c>false</c> 时生效
    /// </summary>
    public int Width { get; set; } = 20;

    /// <summary>
    /// 排序
    /// </summary>
    public int ColumnIndex { get; set; } = 10000;

    /// <summary>
    /// 自动换行
    /// </summary>
    public bool WrapText { get; set; }

    /// <summary>
    /// 是否隐藏列
    /// </summary>
    public bool Hidden { get; set; }

    /// <summary>
    /// 是否图片数据
    /// </summary>
    public bool IsImage { get; set; } = false;
}
