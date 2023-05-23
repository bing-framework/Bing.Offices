namespace Bing.Offices.Exports.Attributes
{
    /// <summary>
    /// 字符格式化 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class StringFormatterAttribute : Attribute
    {
        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 初始化一个<see cref="StringFormatterAttribute"/>类型的实例
        /// </summary>
        /// <param name="format">格式化</param>
        public StringFormatterAttribute(string format) => Format = format;
    }
}
