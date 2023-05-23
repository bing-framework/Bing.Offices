namespace Bing.Offices.Exports.Attributes
{
    /// <summary>
    /// 行合并 特性（根据指定关键字合并）
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
    public class RowMergedAttribute : Attribute
    {
    }
}
