using System.Collections.Generic;

namespace Bing.Offices.Configurations;

using Bing.Offices.Imports;

/// <summary>
/// 规范化映射文档中的动态列描述，不包含 CLR 类型对象。
/// </summary>
public sealed class ExcelMappingDynamicColumnConfiguration
{
    /// <summary>获取或设置稳定列键。</summary>
    public string Key { get; set; }
    /// <summary>获取或设置展示标题。</summary>
    public string Title { get; set; }
    /// <summary>获取或设置标题别名。</summary>
    public List<string> Aliases { get; set; } = new List<string>();
    /// <summary>获取或设置批准的数据类型名称。</summary>
    public string DataTypeName { get; set; }
    /// <summary>获取或设置同层排序值。</summary>
    public int Order { get; set; }
    /// <summary>获取或设置值转换器名称。</summary>
    public string ConverterName { get; set; }
    /// <summary>获取或设置命名校验器名称。</summary>
    public string ValidatorName { get; set; }
    /// <summary>获取或设置按顺序执行的命名校验规则名称。</summary>
    public List<string> ValidationRuleNames { get; set; } = new List<string>();
    /// <summary>获取或设置按顺序执行的内置校验规则描述。</summary>
    public List<ExcelMappingDynamicValidationConfiguration> ValidationRules { get; set; } =
        new List<ExcelMappingDynamicValidationConfiguration>();
    /// <summary>获取或设置数字格式。</summary>
    public string NumberFormat { get; set; }
    /// <summary>获取或设置物理列索引。</summary>
    public int? ColumnIndex { get; set; }
    /// <summary>获取或设置相对布局键，格式为 before:列键 或 after:列键。</summary>
    public string PlacementKey { get; set; }
    /// <summary>获取或设置图片多值策略。</summary>
    public ExcelImageMultiplicityPolicy ImageMultiplicity { get; set; } = ExcelImageMultiplicityPolicy.First;
}
