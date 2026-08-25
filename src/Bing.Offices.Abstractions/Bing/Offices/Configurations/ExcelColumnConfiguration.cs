namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 属性列的请求级配置。
/// </summary>
public sealed class ExcelColumnConfiguration
{
    /// <summary>
    /// 获取或设置实体属性名称。
    /// </summary>
    public string PropertyName { get; set; }

    /// <summary>
    /// 获取或设置列标题。
    /// </summary>
    public string Title { get; set; }

    /// <summary>
    /// 获取或设置是否显式清除低优先级列标题。
    /// </summary>
    public bool ClearTitle { get; set; }

    /// <summary>
    /// 获取或设置历史表头别名集合。
    /// </summary>
    public List<string> Aliases { get; set; } = new List<string>();

    /// <summary>
    /// 获取或设置是否显式清空低优先级表头别名。
    /// </summary>
    public bool ClearAliases { get; set; }

    /// <summary>
    /// 获取或设置从零开始的导出列索引。
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// 获取或设置是否显式重置低优先级列索引。
    /// </summary>
    public bool ResetColumnIndex { get; set; }

    /// <summary>
    /// 获取或设置是否忽略该属性。
    /// </summary>
    public bool? Ignored { get; set; }

    /// <summary>
    /// 获取或设置是否显式重置低优先级忽略标记。
    /// </summary>
    public bool ResetIgnored { get; set; }

    /// <summary>
    /// 获取或设置单元格格式化字符串。
    /// </summary>
    public string Formatter { get; set; }

    /// <summary>
    /// 获取或设置是否显式清除低优先级格式化字符串。
    /// </summary>
    public bool ClearFormatter { get; set; }

    /// <summary>
    /// 获取或设置小数精度。
    /// </summary>
    public byte? DecimalScale { get; set; }

    /// <summary>
    /// 获取或设置是否显式重置低优先级小数精度。
    /// </summary>
    public bool ResetDecimalScale { get; set; }

    /// <summary>
    /// 获取或设置已注册值转换器的名称。
    /// </summary>
    public string ConverterName { get; set; }

    /// <summary>
    /// 获取或设置是否显式清除低优先级转换器名称。
    /// </summary>
    public bool ClearConverterName { get; set; }

    /// <summary>
    /// 获取或设置导入文本空白策略。
    /// </summary>
    public Imports.ExcelWhitespacePolicy? ImportWhitespace { get; set; }

    /// <summary>
    /// 获取或设置是否显式重置低优先级空白策略。
    /// </summary>
    public bool ResetImportWhitespace { get; set; }

    /// <summary>
    /// 获取或设置已注册校验规则的名称集合。
    /// </summary>
    public List<string> ValidationRuleNames { get; set; } = new List<string>();

    /// <summary>
    /// 获取或设置需要从低优先级配置移除的校验规则名称。
    /// </summary>
    public List<string> ValidationRuleNamesToRemove { get; set; } = new List<string>();

    /// <summary>
    /// 获取或设置是否清空低优先级校验规则。
    /// </summary>
    public bool ClearValidationRules { get; set; }

    /// <summary>
    /// 获取或设置校验规则集合的合并方式。
    /// </summary>
    public ExcelValidationRuleMergeMode? ValidationRuleMergeMode { get; set; }

    /// <summary>
    /// 获取或设置显示值映射集合。
    /// </summary>
    public List<ExcelValueMappingConfiguration> ValueMappings { get; set; } = new List<ExcelValueMappingConfiguration>();

    /// <summary>
    /// 获取或设置是否显式清空低优先级值映射。
    /// </summary>
    public bool ClearValueMappings { get; set; }

    /// <summary>
    /// 获取或设置显示值映射集合的合并方式。
    /// </summary>
    public ExcelValueMappingMergeMode? ValueMappingMergeMode { get; set; }

    /// <summary>
    /// 图片列出现多个图片时的处理策略。
    /// </summary>
    public Imports.ExcelImageMultiplicityPolicy? ImageMultiplicity { get; set; }

    /// <summary>
    /// 获取或设置是否显式重置低优先级图片多重性策略。
    /// </summary>
    public bool ResetImageMultiplicity { get; set; }
}
