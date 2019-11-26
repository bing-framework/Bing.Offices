using System;

namespace Bing.Offices.Excel.Configurations
{
    /// <summary>
    /// 属性配置
    /// </summary>
    public interface IPropertyConfiguration
    {
    }

    /// <summary>
    /// 属性配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyConfiguration<out TEntity, TProperty> : IPropertyConfiguration
    {
        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index);

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width);

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title);

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter);

        /// <summary>
        /// 设置忽略该属性
        /// </summary>
        IPropertyConfiguration<TEntity, TProperty> Ignored();

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc);

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc);
    }
}
