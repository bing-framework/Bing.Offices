using System.Collections.Generic;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 属性扩展设置
    /// </summary>
    public interface IPropertyExtendFluent
    {
    }

    /// <summary>
    /// 属性扩展设置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyExtendFluent<out TEntity, TProperty> : IPropertyExtendFluent
    {
        /// <summary>
        /// 设置保留小数位数
        /// </summary>
        /// <param name="scale">保留小数位数</param>
        IPropertyExtendFluent<TEntity, TProperty> HasDecimalScale(byte scale = 2);

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="defaultValue">默认值</param>
        IPropertyExtendFluent<TEntity, TProperty> HasDefaultValue(object defaultValue);

        /// <summary>
        /// 设置动态列
        /// </summary>
        /// <param name="columns">动态列</param>
        IPropertyExtendFluent<TEntity, TProperty> HasDynamicColumn(IEnumerable<string> columns);
    }
}
