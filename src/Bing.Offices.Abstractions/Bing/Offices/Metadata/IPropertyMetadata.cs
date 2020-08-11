using System;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义属性元数据
    /// </summary>
    public interface IPropertyMetadata
    {
        ///// <summary>
        ///// 属性名称
        ///// </summary>
        //string PropertyName { get; set; }

        ///// <summary>
        ///// 属性类型
        ///// </summary>
        //Type PropertyType { get; set; }

        ///// <summary>
        ///// 是否忽略
        ///// </summary>
        //bool IsIgnored { get; set; }

        ///// <summary>
        ///// 是否自动清空空格
        ///// </summary>
        //bool IsAutoTrim { get; set; }

        ///// <summary>
        ///// 描述
        ///// </summary>
        //string Description { get; set; }

        ///// <summary>
        ///// 数据列元数据
        ///// </summary>
        //IColumnMetadata Column { get; set; }
    }

    /// <summary>
    /// 定义属性元数据
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyMetadata<out TEntity, TProperty> : IPropertyMetadata
    {
        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        IPropertyMetadata<TEntity, TProperty> HasColumnIndex(int index);

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        IPropertyMetadata<TEntity, TProperty> HasColumnWidth(int width);

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        IPropertyMetadata<TEntity, TProperty> HasColumnTitle(string title);

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        IPropertyMetadata<TEntity, TProperty> HasColumnFormatter(string formatter);

        /// <summary>
        /// 忽略当前列
        /// </summary>
        /// <param name="ignored">是否忽略</param>
        IPropertyMetadata<TEntity, TProperty> Ignored(bool ignored = true);

        /// <summary>
        /// 设置列输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyMetadata<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc);

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyMetadata<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc);

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        IPropertyMetadata<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc);

        /// <summary>
        /// 设置属性扩展
        /// </summary>
        /// <param name="action">操作</param>
        IPropertyMetadata<TEntity, TProperty> HasPropertyExtend(Action<IPropertyExtendMetadata<TEntity, TProperty>> action);

        /// <summary>
        /// 获取属性扩展
        /// </summary>
        IPropertyExtendMetadata<TEntity, TProperty> PropertyExtend();
    }
}
