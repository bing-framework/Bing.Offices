using System.Linq;

namespace Bing.Offices.Configurations;

internal static class MappingConfigurationCloner
{
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

    internal static ExcelMappingLayoutConfiguration Clone(ExcelMappingLayoutConfiguration layout) =>
        layout == null ? null : new ExcelMappingLayoutConfiguration
        {
            ColumnIndex = layout.ColumnIndex,
            ResetColumnIndex = layout.ResetColumnIndex,
            PlacementKey = layout.PlacementKey,
            ClearPlacementKey = layout.ClearPlacementKey
        };
}
