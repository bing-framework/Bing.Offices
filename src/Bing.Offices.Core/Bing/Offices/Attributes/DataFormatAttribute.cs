namespace Bing.Offices.Attributes;

/// <summary>
/// 指定 Excel 单元格使用的内置或自定义数据格式。
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public class DataFormatAttribute : Attribute
{
    /// <summary>
    /// 获取或设置内置格式。
    /// </summary>
    /// <remarks>参考：https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html</remarks>
    internal short BuiltinFormat { get; set; }

    /// <summary>
    /// 获取或设置自定义格式。
    /// </summary>
    /// <remarks>参考：https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4</remarks>
    public string CustomFormat { get; set; }

    /// <summary>
    /// 初始化一个 <see cref="DataFormatAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="format">Excel 内置数据格式的索引。</param>
    private DataFormatAttribute(short format) => BuiltinFormat = format;

    /// <summary>
    /// 初始化一个 <see cref="DataFormatAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="format">Excel 自定义数据格式字符串。</param>
    public DataFormatAttribute(string format) => CustomFormat = format;
}
