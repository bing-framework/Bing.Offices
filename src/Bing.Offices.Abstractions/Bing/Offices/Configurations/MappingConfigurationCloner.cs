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
            Aliases = (column.Aliases ?? new List<string>()).ToList(),
            ColumnIndex = column.ColumnIndex,
            Ignored = column.Ignored,
            Formatter = column.Formatter,
            DecimalScale = column.DecimalScale,
            ConverterName = column.ConverterName,
            ImportWhitespace = column.ImportWhitespace,
            ValidationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList(),
            ValidationRuleNamesToRemove = (column.ValidationRuleNamesToRemove ?? new List<string>()).ToList(),
            ClearValidationRules = column.ClearValidationRules,
            ValidationRuleMergeMode = column.ValidationRuleMergeMode,
            ValueMappings = (column.ValueMappings ?? new List<ExcelValueMappingConfiguration>()).Select(mapping =>
                mapping == null ? null : new ExcelValueMappingConfiguration { Text = mapping.Text, Value = mapping.Value })
                .ToList(),
            ValueMappingMergeMode = column.ValueMappingMergeMode,
            ImageMultiplicity = column.ImageMultiplicity
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
            BodyStyleKey = style.BodyStyleKey,
            NumberFormat = style.NumberFormat
        };

    internal static ExcelMappingLayoutConfiguration Clone(ExcelMappingLayoutConfiguration layout) =>
        layout == null ? null : new ExcelMappingLayoutConfiguration
        {
            ColumnIndex = layout.ColumnIndex,
            PlacementKey = layout.PlacementKey
        };
}
