using System.Collections.Generic;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 属性扩展元数据
    /// </summary>
    public interface IPropertyExtendMetadata
    {
    }

    /// <summary>
    /// 属性扩展元数据
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyExtendMetadata<out TEntity, TProperty> : IPropertyExtendMetadata
    {
        /// <summary>
        /// 设置保留小数位数
        /// </summary>
        /// <param name="scale">保留小数位数</param>
        IPropertyExtendMetadata<TEntity, TProperty> HasDecimalScale(byte scale = 2);

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="defaultValue">默认值</param>
        IPropertyExtendMetadata<TEntity, TProperty> HasDefaultValue(object defaultValue);

        /// <summary>
        /// 设置动态列
        /// </summary>
        /// <param name="columns">动态列</param>
        IPropertyExtendMetadata<TEntity, TProperty> HasDynamicColumn(IEnumerable<string> columns);
    }
}
