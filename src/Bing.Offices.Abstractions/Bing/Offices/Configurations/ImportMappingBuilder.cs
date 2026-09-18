using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// 导入方向映射构建器。
/// </summary>
/// <typeparam name="T">导入模型类型。</typeparam>
public sealed class ImportMappingBuilder<T> where T : class, new()
{
    /// <summary>
    /// 保存当前导入方向尚未构建的可变列配置。
    /// </summary>
    private readonly ExcelMappingConfiguration _configuration = new();

    /// <summary>
    /// 配置导入模型属性。
    /// </summary>
    /// <typeparam name="TProperty">所配置属性的类型。</typeparam>
    /// <param name="expression">指向导入模型直接属性的表达式。</param>
    /// <returns>指定属性的导入映射配置器。</returns>
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
    /// <param name="sourceKind">当前配置快照的来源类型。</param>
    /// <returns>当前导入方向的独立配置快照。</returns>
    public ExcelMappingConfiguration Build(MappingSourceKind sourceKind = MappingSourceKind.Profile) =>
        MappingConfigurationCloner.Clone(_configuration, sourceKind);

    /// <summary>
    /// 从表达式解析导入模型的直接属性。
    /// </summary>
    /// <typeparam name="TProperty">属性值类型。</typeparam>
    /// <param name="expression">指向导入模型直接属性的表达式。</param>
    /// <returns>表达式指向的属性信息。</returns>
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
/// <typeparam name="T">导入模型类型。</typeparam>
/// <typeparam name="TProperty">当前属性的类型。</typeparam>
public sealed class ImportColumnMappingBuilder<T, TProperty> where T : class, new()
{
    /// <summary>
    /// 保存当前属性配置器所属的导入构建器。
    /// </summary>
    private readonly ImportMappingBuilder<T> _owner;
    /// <summary>
    /// 当前属性对应的可变列配置。
    /// </summary>
    private readonly ExcelColumnConfiguration _configuration;

    /// <summary>
    /// 初始化一个 <see cref="ImportColumnMappingBuilder{T,TProperty}" /> 类型的实例。
    /// </summary>
    /// <param name="owner">当前属性配置器所属的导入构建器。</param>
    /// <param name="configuration">当前属性对应的可变列配置。</param>
    internal ImportColumnMappingBuilder(ImportMappingBuilder<T> owner, ExcelColumnConfiguration configuration)
    {
        _owner = owner;
        _configuration = configuration;
    }

    /// <summary>
    /// 设置导入表头。
    /// </summary>
    /// <param name="header">导入文件中匹配的表头文本。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasHeader(string header)
    {
        _configuration.Title = header;
        return this;
    }

    /// <summary>
    /// 添加导入表头别名。
    /// </summary>
    /// <param name="aliases">可接受的历史或替代表头文本。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
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
    /// <param name="columnIndex">导入列的零基索引。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasColumnIndex(int columnIndex)
    {
        _configuration.ColumnIndex = columnIndex;
        return this;
    }

    /// <summary>
    /// 设置导入值转换器名称。
    /// </summary>
    /// <param name="converterName">要绑定的命名转换器名称。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasConverter(string converterName)
    {
        _configuration.ConverterName = converterName;
        return this;
    }

    /// <summary>
    /// 设置导入文本空白策略。
    /// </summary>
    /// <param name="policy">导入文本空白的规范化策略。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasWhitespace(Imports.ExcelWhitespacePolicy policy)
    {
        _configuration.ImportWhitespace = policy;
        return this;
    }

    /// <summary>
    /// 设置是否忽略导入属性。
    /// </summary>
    /// <param name="ignored">为 <see langword="true"/> 时跳过该属性，为 <see langword="false"/> 时参与导入。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> Ignored(bool ignored = true)
    {
        _configuration.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 设置图片列多重性策略。
    /// </summary>
    /// <param name="policy">同一单元格包含多个图片时采用的策略。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasImageMultiplicity(Imports.ExcelImageMultiplicityPolicy policy)
    {
        _configuration.ImageMultiplicity = policy;
        return this;
    }

    /// <summary>
    /// 追加命名校验规则。
    /// </summary>
    /// <param name="ruleName">要追加的命名校验规则名称。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> HasValidationRule(string ruleName)
    {
        _configuration.ValidationRuleMergeMode = ExcelValidationRuleMergeMode.Append;
        _configuration.ValidationRuleNames.Add(ruleName);
        return this;
    }

    /// <summary>
    /// 移除指定命名校验规则。
    /// </summary>
    /// <param name="ruleName">要从继承配置中移除的规则名称。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> RemoveValidation(string ruleName)
    {
        _configuration.ValidationRuleNamesToRemove.Add(ruleName);
        return this;
    }

    /// <summary>
    /// 清空低优先级命名校验规则。
    /// </summary>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> ClearValidations()
    {
        _configuration.ClearValidationRules = true;
        return this;
    }

    /// <summary>
    /// 设置显示文本到导入值的映射。
    /// </summary>
    /// <param name="text">导入文件中出现的显示文本。</param>
    /// <param name="value">显示文本对应的属性值。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
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
    /// 追加显示值映射。
    /// </summary>
    /// <remarks>该映射追加到低优先级映射，而不是替换它们。</remarks>
    /// <param name="text">导入文件中出现的显示文本。</param>
    /// <param name="value">显示文本对应的属性值。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ImportColumnMappingBuilder<T, TProperty> AppendMap(string text, TProperty value)
    {
        _configuration.ValueMappingMergeMode = ExcelValueMappingMergeMode.Append;
        return Map(text, value);
    }

    /// <summary>
    /// 返回当前导入方向构建器。
    /// </summary>
    /// <returns>拥有当前属性配置的导入映射构建器。</returns>
    public ImportMappingBuilder<T> And() => _owner;
}
