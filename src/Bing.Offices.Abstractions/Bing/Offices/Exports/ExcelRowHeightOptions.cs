namespace Bing.Offices.Exports;

using System;

/// <summary>
/// 工作表行高配置，单位为 Excel point。
/// </summary>
public sealed class ExcelRowHeightOptions
{
    /// <summary>
    /// Excel 工作表允许的单行最大高度，单位为 point。
    /// </summary>
    private const double MaxHeight = 409.5;

    /// <summary>
    /// 获取或初始化表头行高度；null 表示不覆盖现有行高。
    /// </summary>
    public double? HeaderHeight { get; init; }

    /// <summary>
    /// 获取或初始化正文行高度；null 表示不覆盖现有行高。
    /// </summary>
    public double? BodyHeight { get; init; }

    /// <summary>
    /// 验证行高配置。
    /// </summary>
    public void Validate()
    {
        ValidateHeight(HeaderHeight, nameof(HeaderHeight));
        ValidateHeight(BodyHeight, nameof(BodyHeight));
    }

    /// <summary>
    /// 验证单项行高是否处于 Excel 支持范围内。
    /// </summary>
    /// <param name="height">待验证的行高。</param>
    /// <param name="parameterName">参数名。</param>
    private static void ValidateHeight(double? height, string parameterName)
    {
        if (!height.HasValue)
            return;
        if (double.IsNaN(height.Value) || double.IsInfinity(height.Value)
            || height.Value <= 0 || height.Value > MaxHeight)
            throw new ArgumentOutOfRangeException(parameterName,
                $"行高必须大于 0 且不超过 {MaxHeight} point。");
    }
}
