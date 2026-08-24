using System;
using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Configurations;

internal static class MappingConfigurationMerger
{
    public static ExcelMappingConfiguration Merge(ExcelMappingConfiguration lower,
        ExcelMappingConfiguration higher, MappingSourceKind sourceKind)
    {
        if (lower == null && higher == null)
            return null;
        var result = MappingConfigurationCloner.Clone(lower ?? new ExcelMappingConfiguration(), sourceKind);
        if (higher == null)
            return result;
        foreach (var higherColumn in higher.Columns ?? new List<ExcelColumnConfiguration>())
        {
            if (higherColumn == null || string.IsNullOrWhiteSpace(higherColumn.PropertyName))
                throw new ArgumentException("映射配置的属性名称不能为空。", nameof(higher));
            var lowerColumn = result.Columns.FirstOrDefault(column =>
                string.Equals(column.PropertyName, higherColumn.PropertyName, StringComparison.OrdinalIgnoreCase));
            if (lowerColumn == null)
            {
                result.Columns.Add(MappingConfigurationCloner.Clone(higherColumn));
                continue;
            }
            MergeColumn(lowerColumn, higherColumn);
        }
        if (higher.DynamicColumns != null && higher.DynamicColumns.Count > 0)
            result.DynamicColumns = higher.DynamicColumns.Select(MappingConfigurationCloner.Clone).ToList();
        if (higher.Style != null)
            result.Style = MappingConfigurationCloner.Clone(higher.Style);
        if (higher.Layout != null)
            result.Layout = MappingConfigurationCloner.Clone(higher.Layout);
        result.SourceKind = sourceKind;
        return result;
    }

    private static void MergeColumn(ExcelColumnConfiguration target, ExcelColumnConfiguration source)
    {
        target.Title = source.Title ?? target.Title;
        target.ColumnIndex = source.ColumnIndex ?? target.ColumnIndex;
        target.Ignored = source.Ignored ?? target.Ignored;
        target.Formatter = source.Formatter ?? target.Formatter;
        target.DecimalScale = source.DecimalScale ?? target.DecimalScale;
        target.ConverterName = source.ConverterName ?? target.ConverterName;
        target.ImportWhitespace = source.ImportWhitespace ?? target.ImportWhitespace;
        target.ImageMultiplicity = source.ImageMultiplicity ?? target.ImageMultiplicity;

        if (source.Aliases != null && source.Aliases.Count > 0)
            target.Aliases = source.Aliases.ToList();

        if (source.ClearValidationRules)
            target.ValidationRuleNames.Clear();
        if (source.ValidationRuleNamesToRemove != null)
        {
            foreach (var rule in source.ValidationRuleNamesToRemove)
                target.ValidationRuleNames.RemoveAll(item => string.Equals(item, rule,
                    StringComparison.OrdinalIgnoreCase));
        }
        if (source.ValidationRuleNames != null && source.ValidationRuleNames.Count > 0)
        {
            if (source.ValidationRuleMergeMode == ExcelValidationRuleMergeMode.Replace)
                target.ValidationRuleNames.Clear();
            foreach (var rule in source.ValidationRuleNames)
            {
                if (!target.ValidationRuleNames.Contains(rule, StringComparer.OrdinalIgnoreCase))
                    target.ValidationRuleNames.Add(rule);
            }
        }

        if (source.ValueMappings != null && source.ValueMappings.Count > 0)
        {
            if (source.ValueMappingMergeMode != ExcelValueMappingMergeMode.Append)
                target.ValueMappings.Clear();
            foreach (var mapping in source.ValueMappings)
            {
                target.ValueMappings.RemoveAll(item => string.Equals(item?.Text, mapping?.Text,
                    StringComparison.Ordinal));
                target.ValueMappings.Add(mapping == null ? null : new ExcelValueMappingConfiguration
                {
                    Text = mapping.Text,
                    Value = mapping.Value
                });
            }
        }
        target.ClearValidationRules = false;
        target.ValidationRuleNamesToRemove.Clear();
        target.ValidationRuleMergeMode = null;
        target.ValueMappingMergeMode = null;
    }
}
