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
