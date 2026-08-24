using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// 导入方向映射构建器。
/// </summary>
/// <typeparam name="T">导入模型类型。</typeparam>
public sealed class ImportMappingBuilder<T> where T : class, new()
{
    private readonly ExcelMappingConfiguration _configuration = new();

    /// <summary>
    /// 配置导入模型属性。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> Property<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        var property = ResolveProperty(expression);
        var configuration = _configuration.Columns.FirstOrDefault(column =>
            string.Equals(column.PropertyName, property.Name, StringComparison.OrdinalIgnoreCase));
        if (configuration == null)
        {
            configuration = new ExcelColumnConfiguration { PropertyName = property.Name };
            _configuration.Columns.Add(configuration);
        }
        return new ImportColumnMappingBuilder<T, TProperty>(this, configuration);
    }

    /// <summary>
    /// 创建当前导入方向的配置快照。
    /// </summary>
    public ExcelMappingConfiguration Build(MappingSourceKind sourceKind = MappingSourceKind.Profile) =>
        MappingConfigurationCloner.Clone(_configuration, sourceKind);

    private static PropertyInfo ResolveProperty<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        if (expression == null)
            throw new ArgumentNullException(nameof(expression));
        if (expression.Body is not MemberExpression { Member: PropertyInfo property }
            || property.DeclaringType != typeof(T))
            throw new ArgumentException("表达式必须指向实体的直接属性。", nameof(expression));
        return property;
    }

}

/// <summary>
/// 单个导入属性的方向专用配置器。
/// </summary>
public sealed class ImportColumnMappingBuilder<T, TProperty> where T : class, new()
{
    private readonly ImportMappingBuilder<T> _owner;
    private readonly ExcelColumnConfiguration _configuration;

    internal ImportColumnMappingBuilder(ImportMappingBuilder<T> owner, ExcelColumnConfiguration configuration)
    {
        _owner = owner;
        _configuration = configuration;
    }

    /// <summary>
    /// 设置导入表头。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasHeader(string header)
    {
        _configuration.Title = header;
        return this;
    }

    /// <summary>
    /// 添加导入表头别名。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasAlias(params string[] aliases)
    {
        if (aliases == null)
            throw new ArgumentNullException(nameof(aliases));
        _configuration.Aliases.AddRange(aliases);
        return this;
    }

    /// <summary>
    /// 设置导入列索引。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasColumnIndex(int columnIndex)
    {
        _configuration.ColumnIndex = columnIndex;
        return this;
    }

    /// <summary>
    /// 设置导入值转换器名称。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasConverter(string converterName)
    {
        _configuration.ConverterName = converterName;
        return this;
    }

    /// <summary>
    /// 设置导入文本空白策略。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasWhitespace(Imports.ExcelWhitespacePolicy policy)
    {
        _configuration.ImportWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置是否忽略导入属性。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> Ignored(bool ignored = true)
    {
        _configuration.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 设置图片列多重性策略。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasImageMultiplicity(Imports.ExcelImageMultiplicityPolicy policy)
    {
        _configuration.ImageMultiplicity = policy;
        return this;
    }

    /// <summary>
    /// 追加命名校验规则。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> HasValidationRule(string ruleName)
    {
        _configuration.ValidationRuleMergeMode = ExcelValidationRuleMergeMode.Append;
        _configuration.ValidationRuleNames.Add(ruleName);
        return this;
    }

    /// <summary>
    /// 移除指定命名校验规则。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> RemoveValidation(string ruleName)
    {
        _configuration.ValidationRuleNamesToRemove.Add(ruleName);
        return this;
    }

    /// <summary>
    /// 清空低优先级命名校验规则。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> ClearValidations()
    {
        _configuration.ClearValidationRules = true;
        return this;
    }

    /// <summary>
    /// 设置显示文本到导入值的映射。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> Map(string text, TProperty value)
    {
        _configuration.ValueMappings.Add(new ExcelValueMappingConfiguration
        {
            Text = text,
            Value = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
        });
        return this;
    }

    /// <summary>
    /// 追加显示值映射，而不是替换低优先级映射。
    /// </summary>
    public ImportColumnMappingBuilder<T, TProperty> AppendMap(string text, TProperty value)
    {
        _configuration.ValueMappingMergeMode = ExcelValueMappingMergeMode.Append;
        return Map(text, value);
    }

    /// <summary>
    /// 返回当前导入方向构建器。
    /// </summary>
    public ImportMappingBuilder<T> And() => _owner;
}
