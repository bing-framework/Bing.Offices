using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// 导出方向映射构建器。
/// </summary>
/// <typeparam name="T">导出模型类型。</typeparam>
public sealed class ExportMappingBuilder<T> where T : class, new()
{
    private readonly ExcelMappingConfiguration _configuration = new();

    /// <summary>
    /// 配置导出模型属性。
    /// </summary>
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
    public ExcelMappingConfiguration Build(MappingSourceKind sourceKind = MappingSourceKind.Profile) =>
        MappingConfigurationCloner.Clone(_configuration, sourceKind);

}

/// <summary>
/// 单个导出属性的方向专用配置器。
/// </summary>
public sealed class ExportColumnMappingBuilder<T, TProperty> where T : class, new()
{
    private readonly ExportMappingBuilder<T> _owner;
    private readonly ExcelColumnConfiguration _configuration;

    internal ExportColumnMappingBuilder(ExportMappingBuilder<T> owner, ExcelColumnConfiguration configuration)
    {
        _owner = owner;
        _configuration = configuration;
    }

    /// <summary>
    /// 设置导出表头。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> HasHeader(string header)
    {
        _configuration.Title = header;
        return this;
    }

    /// <summary>
    /// 设置导出列索引。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> HasColumnIndex(int columnIndex)
    {
        _configuration.ColumnIndex = columnIndex;
        return this;
    }

    /// <summary>
    /// 设置导出格式化字符串。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> HasFormatter(string formatter)
    {
        _configuration.Formatter = formatter;
        return this;
    }

    /// <summary>
    /// 设置导出小数精度。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> HasDecimalScale(byte decimalScale)
    {
        _configuration.DecimalScale = decimalScale;
        return this;
    }

    /// <summary>
    /// 设置是否忽略导出属性。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> Ignored(bool ignored = true)
    {
        _configuration.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 设置图片列多重性策略。
    /// </summary>
    public ExportColumnMappingBuilder<T, TProperty> HasImageMultiplicity(Imports.ExcelImageMultiplicityPolicy policy)
    {
        _configuration.ImageMultiplicity = policy;
        return this;
    }

    /// <summary>
    /// 返回当前导出方向构建器。
    /// </summary>
    public ExportMappingBuilder<T> And() => _owner;
}
