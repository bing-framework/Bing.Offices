using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Collections.ObjectModel;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Validations;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

internal sealed class ExcelDynamicMappingColumn : IExcelDynamicMappingColumn
{
    /// <summary>从动态列配置及已绑定的转换器和校验器创建不可变动态列计划。</summary>
    /// <param name="column">规范化后的动态列配置。</param>
    /// <param name="converters">已绑定的值转换器。</param>
    /// <param name="validations">已绑定的校验规则。</param>
    internal ExcelDynamicMappingColumn(ExcelMappingDynamicColumnConfiguration column,
        IReadOnlyList<IExcelValueConverter> converters, IReadOnlyList<IExcelValidationBinding> validations)
    {
        Key = column.Key;
        Title = column.Title;
        Aliases = new ReadOnlyCollection<string>((column.Aliases ?? new List<string>()).ToArray());
        DataTypeName = column.DataTypeName ?? "string";
        Order = column.Order;
        ConverterName = column.ConverterName;
        ValidatorName = column.ValidatorName;
        var validationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList();
        if (!string.IsNullOrWhiteSpace(column.ValidatorName)
            && !validationRuleNames.Contains(column.ValidatorName, StringComparer.OrdinalIgnoreCase))
            validationRuleNames.Insert(0, column.ValidatorName);
        ValidationRuleNames = new ReadOnlyCollection<string>(validationRuleNames.ToArray());
        ValidationRules = new ReadOnlyCollection<ExcelMappingDynamicValidationConfiguration>(
            (column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
                .Select(rule => new ExcelMappingDynamicValidationConfiguration
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
                }).ToArray());
        NumberFormat = column.NumberFormat;
        ColumnIndex = column.ColumnIndex;
        PlacementKey = column.PlacementKey;
        ImageMultiplicity = column.ImageMultiplicity;
        ValueConverters = new ReadOnlyCollection<IExcelValueConverter>(converters.ToArray());
        ValidationBindings = new ReadOnlyCollection<IExcelValidationBinding>(validations.ToArray());
        var unique = (column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
            .FirstOrDefault(rule => rule != null && string.Equals(rule.Name?.Trim(), "unique",
                StringComparison.OrdinalIgnoreCase));
        IsUnique = unique != null;
        UniqueIgnoreEmpty = unique?.IgnoreEmpty ?? true;
    }

    /// <inheritdoc />
    public string Key { get; }
    /// <inheritdoc />
    public string Title { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> Aliases { get; }
    /// <inheritdoc />
    public string DataTypeName { get; }
    /// <inheritdoc />
    public int Order { get; }
    /// <inheritdoc />
    public string ConverterName { get; }
    /// <inheritdoc />
    public string ValidatorName { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>获取动态列配置中声明的内置校验规则快照。</summary>
    public IReadOnlyList<ExcelMappingDynamicValidationConfiguration> ValidationRules { get; }
    /// <inheritdoc />
    public string NumberFormat { get; }
    /// <inheritdoc />
    public int? ColumnIndex { get; }
    /// <inheritdoc />
    public string PlacementKey { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <inheritdoc />
    public bool IsUnique { get; }
    /// <inheritdoc />
    public bool UniqueIgnoreEmpty { get; }
}
