using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// 导出方向映射构建器。
/// </summary>
/// <typeparam name="T">导出模型类型。</typeparam>
public sealed class ExportMappingBuilder<T> where T : class, new()
{
    /// <summary>
    /// 保存当前导出方向尚未构建的可变列配置。
    /// </summary>
    private readonly ExcelMappingConfiguration _configuration = new();

    /// <summary>
    /// 配置导出模型属性。
    /// </summary>
    /// <typeparam name="TProperty">所配置属性的类型。</typeparam>
    /// <param name="expression">指向导出模型直接属性的表达式。</param>
    /// <returns>指定属性的导出映射配置器。</returns>
    public ExportColumnMappingBuilder<T, TProperty> Property<TProperty>(Expression<Func<T, TProperty>> expression)
    {
        if (expression == null)
            throw new ArgumentNullException(nameof(expression));
        if (expression.Body is not MemberExpression { Member: PropertyInfo property }
            || property.DeclaringType != typeof(T))
            throw new ArgumentException("表达式必须指向实体的直接属性。", nameof(expression));
        var configuration = _configuration.Columns.FirstOrDefault(column =>
            string.Equals(column.PropertyName, property.Name, StringComparison.OrdinalIgnoreCase));
        if (configuration == null)
        {
            configuration = new ExcelColumnConfiguration { PropertyName = property.Name };
            _configuration.Columns.Add(configuration);
        }
        return new ExportColumnMappingBuilder<T, TProperty>(this, configuration);
    }

    /// <summary>
    /// 创建当前导出方向的配置快照。
    /// </summary>
    /// <param name="sourceKind">当前配置快照的来源类型。</param>
    /// <returns>当前导出方向的独立配置快照。</returns>
    public ExcelMappingConfiguration Build(MappingSourceKind sourceKind = MappingSourceKind.Profile) =>
        MappingConfigurationCloner.Clone(_configuration, sourceKind);

}
