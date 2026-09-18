namespace Bing.Offices.Configurations;

/// <summary>
/// 复制映射配置及其嵌套定义的内部辅助类。
/// </summary>
internal static class MappingConfigurationCloner
{
    /// <summary>
    /// 复制映射配置及其嵌套定义。
    /// </summary>
    /// <param name="configuration">待复制的映射配置。</param>
    /// <param name="sourceKind">副本的来源类型。</param>
    /// <returns>映射配置的独立副本。</returns>
    public static ExcelMappingConfiguration Clone(ExcelMappingConfiguration configuration,
        MappingSourceKind sourceKind)
    {
        if (configuration == null)
            throw new ArgumentNullException(nameof(configuration));
        return new ExcelMappingConfiguration
        {
            SourceKind = sourceKind,
            Profile = configuration.Profile,
            ModelAlias = configuration.ModelAlias,
            ClearDynamicColumns = configuration.ClearDynamicColumns,
            DynamicColumnKeysToRemove = (configuration.DynamicColumnKeysToRemove ?? new List<string>()).ToList(),
            DynamicColumnMergeMode = configuration.DynamicColumnMergeMode,
            ResetStyle = configuration.ResetStyle,
            ResetLayout = configuration.ResetLayout,
            Columns = (configuration.Columns ?? new List<ExcelColumnConfiguration>()).Select(Clone).ToList(),
            DynamicColumns = (configuration.DynamicColumns ?? new List<ExcelMappingDynamicColumnConfiguration>())
                .Select(Clone).ToList(),
            Style = Clone(configuration.Style),
            Layout = Clone(configuration.Layout)
        };
    }

    /// <summary>
    /// 复制单个列映射配置。
    /// </summary>
    /// <param name="column">待复制的列映射配置。</param>
    /// <returns>列映射配置的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    public static ExcelColumnConfiguration Clone(ExcelColumnConfiguration column)
    {
        if (column == null)
            return null;
        return new ExcelColumnConfiguration
        {
            PropertyName = column.PropertyName,
            Title = column.Title,
            ClearTitle = column.ClearTitle,
            Aliases = (column.Aliases ?? new List<string>()).ToList(),
            ClearAliases = column.ClearAliases,
            ColumnIndex = column.ColumnIndex,
            ResetColumnIndex = column.ResetColumnIndex,
            Ignored = column.Ignored,
            ResetIgnored = column.ResetIgnored,
            Formatter = column.Formatter,
            ClearFormatter = column.ClearFormatter,
            DecimalScale = column.DecimalScale,
            ResetDecimalScale = column.ResetDecimalScale,
            ConverterName = column.ConverterName,
            ClearConverterName = column.ClearConverterName,
            ImportWhitespace = column.ImportWhitespace,
            ResetImportWhitespace = column.ResetImportWhitespace,
            ValidationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList(),
            ValidationRuleNamesToRemove = (column.ValidationRuleNamesToRemove ?? new List<string>()).ToList(),
            ClearValidationRules = column.ClearValidationRules,
            ValidationRuleMergeMode = column.ValidationRuleMergeMode,
            ValueMappings = (column.ValueMappings ?? new List<ExcelValueMappingConfiguration>()).Select(mapping =>
                mapping == null ? null : new ExcelValueMappingConfiguration { Text = mapping.Text, Value = mapping.Value })
                .ToList(),
            ClearValueMappings = column.ClearValueMappings,
            ValueMappingMergeMode = column.ValueMappingMergeMode,
            ImageMultiplicity = column.ImageMultiplicity,
            ResetImageMultiplicity = column.ResetImageMultiplicity
        };
    }

    /// <summary>
    /// 复制动态列映射配置。
    /// </summary>
    /// <param name="column">待复制的动态列映射配置。</param>
    /// <returns>动态列映射配置的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    internal static ExcelMappingDynamicColumnConfiguration Clone(ExcelMappingDynamicColumnConfiguration column) =>
        column == null ? null : new ExcelMappingDynamicColumnConfiguration
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = (column.Aliases ?? new List<string>()).ToList(),
            DataTypeName = column.DataTypeName,
            Order = column.Order,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList(),
            ValidationRules = (column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
                .Select(rule => rule == null ? null : new ExcelMappingDynamicValidationConfiguration
                {
                    Name = rule.Name,
                    Pattern = rule.Pattern,
                    Format = rule.Format,
                    CultureName = rule.CultureName,
                    Min = rule.Min,
                    Max = rule.Max,
                    MaxValue = rule.MaxValue,
                    MaxLength = rule.MaxLength,
                    IgnoreEmpty = rule.IgnoreEmpty
                }).ToList(),
            NumberFormat = column.NumberFormat,
            ColumnIndex = column.ColumnIndex,
            PlacementKey = column.PlacementKey,
            ImageMultiplicity = column.ImageMultiplicity
        };

    /// <summary>
    /// 复制映射样式配置。
    /// </summary>
    /// <param name="style">待复制的样式配置。</param>
    /// <returns>样式配置的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    internal static ExcelMappingStyleConfiguration Clone(ExcelMappingStyleConfiguration style) =>
        style == null ? null : new ExcelMappingStyleConfiguration
        {
            HeaderStyleKey = style.HeaderStyleKey,
            ClearHeaderStyleKey = style.ClearHeaderStyleKey,
            BodyStyleKey = style.BodyStyleKey,
            ClearBodyStyleKey = style.ClearBodyStyleKey,
            NumberFormat = style.NumberFormat,
            ClearNumberFormat = style.ClearNumberFormat
        };

    /// <summary>
    /// 复制映射布局配置。
    /// </summary>
    /// <param name="layout">待复制的布局配置。</param>
    /// <returns>布局配置的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    internal static ExcelMappingLayoutConfiguration Clone(ExcelMappingLayoutConfiguration layout) =>
        layout == null ? null : new ExcelMappingLayoutConfiguration
        {
            ColumnIndex = layout.ColumnIndex,
            ResetColumnIndex = layout.ResetColumnIndex,
            PlacementKey = layout.PlacementKey,
            ClearPlacementKey = layout.ClearPlacementKey
        };
}
