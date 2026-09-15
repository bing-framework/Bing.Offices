using System.Linq.Expressions;
using System.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// 导出方向映射构建器。
/// </summary>
/// <typeparam name="T">导出模型类型。</typeparam>
public sealed class ExportMappingBuilder<T> where T : class, new()
{
    /// <summary>保存当前导出方向尚未构建的可变列配置。</summary>
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

/// <summary>
/// 单个导出属性的方向专用配置器。
/// </summary>
/// <typeparam name="T">导出模型类型。</typeparam>
/// <typeparam name="TProperty">当前属性的类型。</typeparam>
public sealed class ExportColumnMappingBuilder<T, TProperty> where T : class, new()
{
    /// <summary>返回当前属性配置器所属的导出构建器。</summary>
    private readonly ExportMappingBuilder<T> _owner;
    /// <summary>当前属性对应的可变列配置。</summary>
    private readonly ExcelColumnConfiguration _configuration;

    /// <summary>初始化一个 <see cref="ExportColumnMappingBuilder{T,TProperty}" /> 类型的实例。</summary>
    /// <param name="owner">当前属性配置器所属的导出构建器。</param>
    /// <param name="configuration">当前属性对应的可变列配置。</param>
    internal ExportColumnMappingBuilder(ExportMappingBuilder<T> owner, ExcelColumnConfiguration configuration)
    {
        _owner = owner;
        _configuration = configuration;
    }

    /// <summary>
    /// 设置导出表头。
    /// </summary>
    /// <param name="header">导出文件中使用的表头文本。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasHeader(string header)
    {
        _configuration.Title = header;
        return this;
    }

    /// <summary>
    /// 设置导出列索引。
    /// </summary>
    /// <param name="columnIndex">导出列的零基索引。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasColumnIndex(int columnIndex)
    {
        _configuration.ColumnIndex = columnIndex;
        return this;
    }

    /// <summary>
    /// 设置导出格式化字符串。
    /// </summary>
    /// <param name="formatter">用于格式化属性值的格式字符串。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasFormatter(string formatter)
    {
        _configuration.Formatter = formatter;
        return this;
    }

    /// <summary>
    /// 设置导出小数精度。
    /// </summary>
    /// <param name="decimalScale">导出数值保留的小数位数。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasDecimalScale(byte decimalScale)
    {
        _configuration.DecimalScale = decimalScale;
        return this;
    }

    /// <summary>设置导出值转换器名称。</summary>
    /// <param name="converterName">要绑定的命名转换器名称。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasConverter(string converterName)
    {
        _configuration.ConverterName = converterName;
        return this;
    }

    /// <summary>设置属性值到导出显示文本的映射。</summary>
    /// <param name="text">导出文件中显示的文本。</param>
    /// <param name="value">需要映射的属性值。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> Map(string text, TProperty value)
    {
        _configuration.ValueMappings.Add(new ExcelValueMappingConfiguration
        {
            Text = text,
            Value = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
        });
        return this;
    }

    /// <summary>
    /// 设置是否忽略导出属性。
    /// </summary>
    /// <param name="ignored">为 <see langword="true"/> 时跳过该属性，为 <see langword="false"/> 时参与导出。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> Ignored(bool ignored = true)
    {
        _configuration.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 设置图片列多重性策略。
    /// </summary>
    /// <param name="policy">同一单元格包含多个图片时采用的策略。</param>
    /// <returns>当前属性配置器，用于继续配置。</returns>
    public ExportColumnMappingBuilder<T, TProperty> HasImageMultiplicity(Imports.ExcelImageMultiplicityPolicy policy)
    {
        _configuration.ImageMultiplicity = policy;
        return this;
    }

    /// <summary>
    /// 返回当前导出方向构建器。
    /// </summary>
    /// <returns>拥有当前属性配置的导出映射构建器。</returns>
    public ExportMappingBuilder<T> And() => _owner;
}
