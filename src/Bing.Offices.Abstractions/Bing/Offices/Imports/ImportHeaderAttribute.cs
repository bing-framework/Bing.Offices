using System;

namespace Bing.Offices.Imports;

/// <summary>
/// 导入头 特性
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ImportHeaderAttribute : Attribute
{
    /// <summary>
    /// 显示名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 批注
    /// </summary>
    public string Description { get; set; }

    /// <summary>
    /// 自动过滤空格，默认：启用
    /// </summary>
    public bool AutoTrim { get; set; } = true;

    /// <summary>
    /// 处理掉所有的空格，包括中间空格
    /// </summary>
    public bool FixAllSpace { get; set; }

    /// <summary>
    /// 格式化（仅用于模板生成）
    /// </summary>
    public string Format { get; set; }

    /// <summary>
    /// 列索引，如果为0则自动计算
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 是否允许重复
    /// </summary>
    public bool IsAllowRepeat { get; set; } = true;

    /// <summary>
    /// 是否忽略
    /// </summary>
    public bool IsIgnore { get; set; }

    /// <summary>
    /// 是否启用Excel数据验证
    /// </summary>
    /// <remarks>
    /// 对于Excel数据验证，仅用于生成导入模板特性中，作为限制用户对Excel模板数据的约束性
    /// </remarks>
    public bool IsInterValidation { get; set; }

    /// <summary>
    /// 选定单元格时，显示输入的信息
    /// </summary>
    /// <remarks>
    /// 仅在 <see cref="IsInterValidation"/> 启用的情况下
    /// </remarks>
    public string ShowInputMessage { get; set; }

    /// <summary>
    /// 是否图片数据
    /// </summary>
    public bool IsImage { get; set; } = false;
}