using System;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 属性设置
    /// </summary>
    public interface IPropertyFluent
    {
    }

    /// <summary>
    /// 属性设置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyFluent<out TEntity, TProperty> : IPropertyFluent
    {
        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        IPropertyFluent<TEntity, TProperty> HasColumnIndex(int index);

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        IPropertyFluent<TEntity, TProperty> HasColumnWidth(int width);

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        IPropertyFluent<TEntity, TProperty> HasColumnTitle(string title);

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        IPropertyFluent<TEntity, TProperty> HasColumnFormatter(string formatter);

        /// <summary>
        /// 忽略当前列
        /// </summary>
        /// <param name="ignored">是否忽略</param>
        IPropertyFluent<TEntity, TProperty> Ignored(bool ignored = true);

        /// <summary>
        /// 设置列输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyFluent<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc);

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyFluent<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc);

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyFluent<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc);

        /// <summary>
        /// 设置属性扩展
        /// </summary>
        /// <param name="action">操作</param>
        IPropertyFluent<TEntity, TProperty> HasPropertyExtend(Action<IPropertyExtendFluent<TEntity, TProperty>> action);

        /// <summary>
        /// 获取属性扩展
        /// </summary>
        IPropertyExtendFluent<TEntity, TProperty> PropertyExtend();
    }
}
