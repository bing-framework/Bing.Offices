using System;
using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Configurations;

public static class MappingConfigurationMerger
{
    /// <summary>
    /// 将高优先级映射配置合并到低优先级配置上。
    /// </summary>
    public static ExcelMappingConfiguration Merge(ExcelMappingConfiguration lower,
        ExcelMappingConfiguration higher, MappingSourceKind sourceKind)
    {
        if (lower == null && higher == null)
            return null;
        var result = MappingConfigurationCloner.Clone(lower ?? new ExcelMappingConfiguration(), sourceKind);
        if (higher == null)
            return result;
        result.Profile = higher.Profile ?? result.Profile;
        result.ModelAlias = higher.ModelAlias ?? result.ModelAlias;
        if (higher.ClearDynamicColumns)
            result.DynamicColumns.Clear();
        foreach (var key in higher.DynamicColumnKeysToRemove ?? new List<string>())
        {
            if (string.IsNullOrWhiteSpace(key))
                throw new ArgumentException("动态列移除 Key 不能为空。", nameof(higher));
            result.DynamicColumns.RemoveAll(column => string.Equals(column?.Key, key,
                StringComparison.OrdinalIgnoreCase));
        }
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
        {
            if (higher.DynamicColumnMergeMode != ExcelDynamicColumnMergeMode.Append)
                result.DynamicColumns.Clear();
            foreach (var higherDynamicColumn in higher.DynamicColumns)
            {
                if (higherDynamicColumn == null || string.IsNullOrWhiteSpace(higherDynamicColumn.Key))
                    throw new ArgumentException("动态列 Key 不能为空。", nameof(higher));
                var index = result.DynamicColumns.FindIndex(column =>
                    string.Equals(column?.Key, higherDynamicColumn.Key, StringComparison.OrdinalIgnoreCase));
                var clone = MappingConfigurationCloner.Clone(higherDynamicColumn);
                if (index < 0)
                    result.DynamicColumns.Add(clone);
                else
                    result.DynamicColumns[index] = clone;
            }
        }
        result.ClearDynamicColumns = false;
        result.DynamicColumnKeysToRemove.Clear();
        result.DynamicColumnMergeMode = null;
        if (higher.ResetStyle)
            result.Style = null;
        else if (higher.Style != null)
            result.Style = MergeStyle(result.Style, higher.Style);
        result.ResetStyle = false;
        if (higher.ResetLayout)
            result.Layout = null;
        else if (higher.Layout != null)
            result.Layout = MergeLayout(result.Layout, higher.Layout);
        result.ResetLayout = false;
        result.SourceKind = sourceKind;
        return result;
    }

    private static ExcelMappingStyleConfiguration MergeStyle(ExcelMappingStyleConfiguration lower,
        ExcelMappingStyleConfiguration higher)
    {
        var result = MappingConfigurationCloner.Clone(lower) ?? new ExcelMappingStyleConfiguration();
        result.HeaderStyleKey = higher.ClearHeaderStyleKey ? null : higher.HeaderStyleKey ?? result.HeaderStyleKey;
        result.BodyStyleKey = higher.ClearBodyStyleKey ? null : higher.BodyStyleKey ?? result.BodyStyleKey;
        result.NumberFormat = higher.ClearNumberFormat ? null : higher.NumberFormat ?? result.NumberFormat;
        result.ClearHeaderStyleKey = false;
        result.ClearBodyStyleKey = false;
        result.ClearNumberFormat = false;
        return result.HeaderStyleKey == null && result.BodyStyleKey == null && result.NumberFormat == null
            ? null
            : result;
    }

    private static ExcelMappingLayoutConfiguration MergeLayout(ExcelMappingLayoutConfiguration lower,
        ExcelMappingLayoutConfiguration higher)
    {
        var result = MappingConfigurationCloner.Clone(lower) ?? new ExcelMappingLayoutConfiguration();
        result.ColumnIndex = higher.ResetColumnIndex ? null : higher.ColumnIndex ?? result.ColumnIndex;
        result.PlacementKey = higher.ClearPlacementKey ? null : higher.PlacementKey ?? result.PlacementKey;
        result.ResetColumnIndex = false;
        result.ClearPlacementKey = false;
        return !result.ColumnIndex.HasValue && result.PlacementKey == null ? null : result;
    }

    private static void MergeColumn(ExcelColumnConfiguration target, ExcelColumnConfiguration source)
    {
        if (source.ClearTitle)
        {
            target.Title = null;
            target.ClearTitle = true;
        }
        else if (source.Title != null)
        {
            target.Title = source.Title;
            target.ClearTitle = false;
        }

        if (source.ResetColumnIndex)
        {
            target.ColumnIndex = null;
            target.ResetColumnIndex = true;
        }
        else if (source.ColumnIndex.HasValue)
        {
            target.ColumnIndex = source.ColumnIndex;
            target.ResetColumnIndex = false;
        }

        if (source.ResetIgnored)
        {
            target.Ignored = null;
            target.ResetIgnored = true;
        }
        else if (source.Ignored.HasValue)
        {
            target.Ignored = source.Ignored;
            target.ResetIgnored = false;
        }

        if (source.ClearFormatter)
        {
            target.Formatter = null;
            target.ClearFormatter = true;
        }
        else if (source.Formatter != null)
        {
            target.Formatter = source.Formatter;
            target.ClearFormatter = false;
        }

        if (source.ResetDecimalScale)
        {
            target.DecimalScale = null;
            target.ResetDecimalScale = true;
        }
        else if (source.DecimalScale.HasValue)
        {
            target.DecimalScale = source.DecimalScale;
            target.ResetDecimalScale = false;
        }

        if (source.ClearConverterName)
        {
            target.ConverterName = null;
            target.ClearConverterName = true;
        }
        else if (source.ConverterName != null)
        {
            target.ConverterName = source.ConverterName;
            target.ClearConverterName = false;
        }

        if (source.ResetImportWhitespace)
        {
            target.ImportWhitespace = null;
            target.ResetImportWhitespace = true;
        }
        else if (source.ImportWhitespace.HasValue)
        {
            target.ImportWhitespace = source.ImportWhitespace;
            target.ResetImportWhitespace = false;
        }

        if (source.ResetImageMultiplicity)
        {
            target.ImageMultiplicity = null;
            target.ResetImageMultiplicity = true;
        }
        else if (source.ImageMultiplicity.HasValue)
        {
            target.ImageMultiplicity = source.ImageMultiplicity;
            target.ResetImageMultiplicity = false;
        }

        if (source.ClearAliases)
        {
            target.Aliases.Clear();
            target.ClearAliases = true;
        }
        if (source.Aliases != null && source.Aliases.Count > 0)
        {
            target.Aliases = source.Aliases.ToList();
            target.ClearAliases = false;
        }

        if (source.ClearValidationRules)
        {
            target.ValidationRuleNames.Clear();
            target.ClearValidationRules = true;
        }
        if (source.ValidationRuleNamesToRemove != null)
        {
            foreach (var rule in source.ValidationRuleNamesToRemove)
                target.ValidationRuleNames.RemoveAll(item => string.Equals(item, rule,
                    StringComparison.OrdinalIgnoreCase));
        }
            if (source.ValidationRuleMergeMode == ExcelValidationRuleMergeMode.Replace)
            {
                target.ValidationRuleNames.Clear();
                target.ClearValidationRules = true;
            }
            if (source.ValidationRuleNames != null && source.ValidationRuleNames.Count > 0)
        {
            foreach (var rule in source.ValidationRuleNames)
            {
                if (!target.ValidationRuleNames.Contains(rule, StringComparer.OrdinalIgnoreCase))
                    target.ValidationRuleNames.Add(rule);
            }
        }
            if (source.ValidationRuleMergeMode.HasValue)
                target.ValidationRuleMergeMode = source.ValidationRuleMergeMode;

        if (source.ClearValueMappings)
            {
            target.ValueMappings.Clear();
                target.ClearValueMappings = true;
            }
            if (source.ValueMappingMergeMode == ExcelValueMappingMergeMode.Replace)
            {
                target.ValueMappings.Clear();
                target.ClearValueMappings = true;
            }
            if (source.ValueMappings != null && source.ValueMappings.Count > 0)
        {
                if (source.ValueMappingMergeMode != ExcelValueMappingMergeMode.Append)
                {
                target.ValueMappings.Clear();
                    target.ClearValueMappings = true;
                }
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
        if (source.ValueMappingMergeMode.HasValue)
            target.ValueMappingMergeMode = source.ValueMappingMergeMode;
        target.ValidationRuleNamesToRemove.Clear();
    }
}
