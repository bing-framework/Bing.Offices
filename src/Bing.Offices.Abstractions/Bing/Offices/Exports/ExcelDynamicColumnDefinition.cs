using System;
using System.Collections.Generic;
using Bing.Offices.Imports;

namespace Bing.Offices.Exports;

/// <summary>
/// 请求级不可变动态列定义。
/// </summary>
public sealed class ExcelDynamicColumnDefinition
{
    /// <summary>
    /// 获取或设置稳定列键。导入导出值均使用该键。
    /// </summary>
    public string Key { get; init; }

    /// <summary>
    /// 获取或设置展示标题。
    /// </summary>
    public string Title { get; init; }

    /// <summary>
    /// 获取或设置导入时接受的历史标题别名。
    /// </summary>
    public IReadOnlyList<string> Aliases { get; init; } = Array.Empty<string>();

    /// <summary>
    /// 获取或设置数据类型。
    /// </summary>
    public Type DataType { get; init; } = typeof(string);

    /// <summary>
    /// 获取或设置同一布局层内的排序值。
    /// </summary>
    public int Order { get; init; }

    /// <summary>
    /// 获取或设置相对固定列位置或显式物理列索引。
    /// </summary>
    public ExcelColumnPlacement Placement { get; init; }

    /// <summary>
    /// 获取或设置显式物理列索引。设置后不能再通过 <see cref="Placement"/> 指定相对位置。
    /// </summary>
    public int? PhysicalColumnIndex { get; init; }

    /// <summary>
    /// 获取或设置 Excel 数字格式。
    /// </summary>
    public string NumberFormat { get; init; }

    /// <summary>
    /// 获取或设置动态列表头样式。
    /// </summary>
    public Styles.ExcelCellStyle HeaderStyle { get; init; }

    /// <summary>
    /// 获取或设置动态列正文样式。
    /// </summary>
    public Styles.ExcelCellStyle BodyStyle { get; init; }

    /// <summary>
    /// 获取或设置注册值转换器名称。
    /// </summary>
    public string ConverterName { get; init; }

    /// <summary>
    /// 获取或设置注册校验规则名称。
    /// </summary>
    public string ValidatorName { get; init; }

    /// <summary>
    /// 获取或设置按顺序执行的注册校验规则名称。
    /// </summary>
    public IReadOnlyList<string> ValidationRuleNames { get; init; } = Array.Empty<string>();

    /// <summary>
    /// 图片列出现多个图片时的处理策略。
    /// </summary>
    public ExcelImageMultiplicityPolicy ImageMultiplicity { get; init; } = ExcelImageMultiplicityPolicy.First;
}
