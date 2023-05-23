namespace Bing.Offices.Exports.Attributes
{
    /// <summary>
    /// 表头样式 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class HeaderStyleAttribute : StyleAttribute
    {
        /// <summary>
        /// 初始化一个<see cref="HeaderStyleAttribute"/>类型的实例
        /// </summary>
        /// <param name="isBold">加粗</param>
        public HeaderStyleAttribute(bool isBold) => Style.IsBold = isBold;

        /// <summary>
        /// 初始化一个<see cref="HeaderStyleAttribute"/>类型的实例
        /// </summary>
        public HeaderStyleAttribute() { }
    }
}
