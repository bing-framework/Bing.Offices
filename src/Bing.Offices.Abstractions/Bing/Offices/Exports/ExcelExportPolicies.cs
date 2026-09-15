namespace Bing.Offices.Exports;

using System;

/// <summary>
/// 导出列宽计算模式。
/// </summary>
public enum ExcelColumnWidthMode
{
    /// <summary>不修改列宽。</summary>
    None,
    /// <summary>使用固定字符宽度。</summary>
    Fixed,
    /// <summary>使用提供程序自动计算。</summary>
    AutoFit,
    /// <summary>使用受样本限制的自适应估算。</summary>
    Adaptive
}

/// <summary>
/// 批注冲突处理策略。
/// </summary>
public enum ExcelCommentConflictPolicy
{
    /// <summary>保留已有批注。</summary>
    Preserve,
    /// <summary>追加新批注文本。</summary>
    Append,
    /// <summary>替换已有批注。</summary>
    Replace,
    /// <summary>出现冲突时失败。</summary>
    Fail
}

/// <summary>
/// 模板单元格写入策略。
/// </summary>
public enum ExcelTemplateCellOverwritePolicy
{
    /// <summary>
    /// 写入值时保留模板样式和批注；已有公式会被导出值替换。
    /// </summary>
    PreserveTemplate,
    /// <summary>
    /// 写入值前清除模板样式和批注；已有公式会被导出值替换。
    /// </summary>
    ReplaceTemplate
}

/// <summary>
/// 工作表列宽配置，单位为 Excel 字符宽度。
/// </summary>
public sealed class ExcelColumnWidthOptions
{
    /// <summary>获取或初始化列宽计算模式；默认值为 <see cref="ExcelColumnWidthMode.None"/>。</summary>
    public ExcelColumnWidthMode Mode { get; init; }

    /// <summary>获取或初始化Fixed 模式使用的固定宽度，单位为 Excel 字符；其他模式不使用此值。</summary>
    public double? FixedWidth { get; init; }

    /// <summary>获取或初始化启用列宽计算时的最小宽度，单位为 Excel 字符；为 null 时不设下限。</summary>
    public double? MinWidth { get; init; }

    /// <summary>获取或初始化启用列宽计算时的最大宽度，单位为 Excel 字符；为 null 时按 255 个字符限制。</summary>
    public double? MaxWidth { get; init; }

    /// <summary>获取或初始化Adaptive 模式最多采样的数据行数，默认值为 100；其他模式不使用此值。</summary>
    public int SampleRows { get; init; } = 100;

    /// <summary>验证列宽配置。</summary>
    public void Validate()
    {
        if (!Enum.IsDefined(typeof(ExcelColumnWidthMode), Mode))
            throw new ArgumentOutOfRangeException(nameof(Mode));
        if (FixedWidth <= 0)
            throw new ArgumentOutOfRangeException(nameof(FixedWidth));
        if (MinWidth <= 0)
            throw new ArgumentOutOfRangeException(nameof(MinWidth));
        if (MaxWidth <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxWidth));
        if (MinWidth.HasValue && MaxWidth.HasValue && MinWidth > MaxWidth)
            throw new ArgumentException("列宽最小值不能大于最大值。", nameof(MaxWidth));
        if (SampleRows <= 0)
            throw new ArgumentOutOfRangeException(nameof(SampleRows));
        if (Mode == ExcelColumnWidthMode.Fixed && !FixedWidth.HasValue)
            throw new ArgumentException("Fixed 模式必须提供 FixedWidth。", nameof(FixedWidth));
    }
}
