namespace Bing.Offices.Configurations;

/// <summary>
/// 规范化动态列的内置校验规则描述。
/// </summary>
public sealed class ExcelMappingDynamicValidationConfiguration
{
    /// <summary>获取或设置规则名称：required、regex、date、maxValue、range、maxLength 或 unique。</summary>
    public string Name { get; set; }

    /// <summary>获取或设置正则表达式。</summary>
    public string Pattern { get; set; }

    /// <summary>获取或设置日期格式。</summary>
    public string Format { get; set; }

    /// <summary>获取或设置日期区域性名称。</summary>
    public string CultureName { get; set; }

    /// <summary>获取或设置区间最小值。</summary>
    public double? Min { get; set; }

    /// <summary>获取或设置区间最大值。</summary>
    public double? Max { get; set; }

    /// <summary>获取或设置最大值规则的上限。</summary>
    public double? MaxValue { get; set; }

    /// <summary>获取或设置最大字符数。</summary>
    public int? MaxLength { get; set; }

    /// <summary>获取或设置唯一性规则是否忽略空值。</summary>
    public bool IgnoreEmpty { get; set; } = true;
}
