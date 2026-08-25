using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 请求级映射配置入口。
/// </summary>
public static class ExcelMapping
{
    /// <summary>
    /// 创建指定实体类型的映射配置构建器。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    public static ExcelMappingBuilder<T> For<T>() where T : class, new() => new();
}

/// <summary>
/// Excel 请求级映射配置构建器。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
public sealed class ExcelMappingBuilder<T> where T : class, new()
{
    private readonly ExcelMappingConfiguration _configuration = new();

    /// <summary>
    /// 配置指定实体属性。
    /// </summary>
    /// <typeparam name="TProperty">属性类型。</typeparam>
    /// <param name="expression">属性表达式。</param>
    public ExcelColumnMappingBuilder<T, TProperty> Property<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        if (expression == null)
            throw new ArgumentNullException(nameof(expression));
        if (expression.Body is not MemberExpression { Member: PropertyInfo property } || property.DeclaringType != typeof(T))
            throw new ArgumentException("表达式必须指向实体的直接属性。", nameof(expression));
        var configuration = _configuration.Columns.FirstOrDefault(column =>
            string.Equals(column.PropertyName, property.Name, StringComparison.OrdinalIgnoreCase));
        if (configuration == null)
        {
            configuration = new ExcelColumnConfiguration { PropertyName = property.Name };
            _configuration.Columns.Add(configuration);
        }
        return new ExcelColumnMappingBuilder<T, TProperty>(this, configuration);
    }

    /// <summary>
    /// 创建当前配置的独立快照。
    /// </summary>
    public ExcelMappingConfiguration Build() => new()
    {
        Columns = _configuration.Columns.Select(CloneColumn).ToList()
    };

    private static ExcelColumnConfiguration CloneColumn(ExcelColumnConfiguration source) => new()
    {
        PropertyName = source.PropertyName,
        Title = source.Title,
        Aliases = source.Aliases.ToList(),
        ColumnIndex = source.ColumnIndex,
        Ignored = source.Ignored,
        Formatter = source.Formatter,
        DecimalScale = source.DecimalScale,
        ConverterName = source.ConverterName,
        ImportWhitespace = source.ImportWhitespace,
        ValidationRuleNames = source.ValidationRuleNames.ToList(),
        ValidationRuleNamesToRemove = source.ValidationRuleNamesToRemove.ToList(),
        ClearValidationRules = source.ClearValidationRules,
        ValidationRuleMergeMode = source.ValidationRuleMergeMode,
        ValueMappings = source.ValueMappings.Select(mapping => new ExcelValueMappingConfiguration
        {
            Text = mapping.Text,
            Value = mapping.Value
        }).ToList(),
        ValueMappingMergeMode = source.ValueMappingMergeMode,
        ImageMultiplicity = source.ImageMultiplicity
    };
}

/// <summary>
/// 单个 Excel 属性列的 Fluent 配置器。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
/// <typeparam name="TProperty">属性类型。</typeparam>
public sealed class ExcelColumnMappingBuilder<T, TProperty> where T : class, new()
{
    private readonly ExcelMappingBuilder<T> _owner;
    private readonly ExcelColumnConfiguration _configuration;

    internal ExcelColumnMappingBuilder(ExcelMappingBuilder<T> owner, ExcelColumnConfiguration configuration)
    {
        _owner = owner;
        _configuration = configuration;
    }

    /// <summary>
    /// 设置列标题。
    /// </summary>
    /// <param name="title">列标题。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasTitle(string title)
    {
        _configuration.Title = title;
        return this;
    }

    /// <summary>
    /// 设置从零开始的导出列索引。
    /// </summary>
    /// <param name="columnIndex">列索引。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasColumnIndex(int columnIndex)
    {
        _configuration.ColumnIndex = columnIndex;
        return this;
    }

    /// <summary>
    /// 设置单元格格式化字符串。
    /// </summary>
    /// <param name="formatter">格式化字符串。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasFormatter(string formatter)
    {
        _configuration.Formatter = formatter;
        return this;
    }

    /// <summary>
    /// 设置小数精度。
    /// </summary>
    /// <param name="decimalScale">小数精度。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasDecimalScale(byte decimalScale)
    {
        _configuration.DecimalScale = decimalScale;
        return this;
    }

    /// <summary>
    /// 设置已注册值转换器的名称。
    /// </summary>
    /// <param name="converterName">转换器名称。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasConverter(string converterName)
    {
        _configuration.ConverterName = converterName;
        return this;
    }

    /// <summary>
    /// 添加已注册校验规则的名称。
    /// </summary>
    /// <param name="ruleName">校验规则名称。</param>
    public ExcelColumnMappingBuilder<T, TProperty> HasValidationRule(string ruleName)
    {
        _configuration.ValidationRuleNames.Add(ruleName);
        return this;
    }

    /// <summary>
    /// 设置是否忽略属性。
    /// </summary>
    /// <param name="ignored">是否忽略。</param>
    public ExcelColumnMappingBuilder<T, TProperty> Ignored(bool ignored = true)
    {
        _configuration.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 添加显示文本到属性值的映射。
    /// </summary>
    /// <param name="text">显示文本。</param>
    /// <param name="value">属性值。</param>
    public ExcelColumnMappingBuilder<T, TProperty> Map(string text, TProperty value)
    {
        _configuration.ValueMappings.Add(new ExcelValueMappingConfiguration
        {
            Text = text,
            Value = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
        });
        return this;
    }

    /// <summary>
    /// 返回当前映射配置构建器。
    /// </summary>
    public ExcelMappingBuilder<T> And() => _owner;
}
