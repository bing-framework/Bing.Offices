namespace Bing.Offices.Attributes;

/// <summary>
/// 数据格式化特性
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public class DataFormatAttribute : Attribute
{
    /// <summary>
    /// 内置格式。
    /// 参考：https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html
    /// </summary>
    internal short BuiltinFormat { get; set; }

    /// <summary>
    /// 自定义格式。
    /// 参考：https://support.office.com/en-nz/article/Create-or-delete-a-custom-number-format-78f2a361-936b-4c03-8772-09fab54be7f4
    /// </summary>
    public string CustomFormat { get; set; }

    /// <summary>
    /// 初始化一个<see cref="DataFormatAttribute"/>类型的实例
    /// </summary>
    /// <param name="format">内置格式</param>
    private DataFormatAttribute(short format) => BuiltinFormat = format;

    /// <summary>
    /// 初始化一个<see cref="DataFormatAttribute"/>类型的实例
    /// </summary>
    /// <param name="format">自定义格式</param>
    public DataFormatAttribute(string format) => CustomFormat = format;
}