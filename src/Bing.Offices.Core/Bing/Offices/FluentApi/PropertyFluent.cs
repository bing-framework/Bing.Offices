using System;
using Bing.Offices.Extensions;
using Bing.Offices.Metadata;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 属性设置
    /// </summary>
    internal class PropertyFluent : IPropertyFluent
    {
        /// <summary>
        /// 属性元数据
        /// </summary>
        public PropertyMetadata Metadata { get; protected set; }

        /// <summary>
        /// 属性扩展设置
        /// </summary>
        public IPropertyExtendFluent Extend { get; protected set; }
    }

    /// <summary>
    /// 属性设置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal sealed class PropertyFluent<TEntity, TProperty> : PropertyFluent, IPropertyFluent<TEntity, TProperty>
    {
        /// <summary>
        /// 初始化一个<see cref="PropertyFluent{TEntity,TProperty}"/>类型的实例
        /// </summary>
        /// <param name="metadata">属性元数据</param>
        public PropertyFluent(PropertyMetadata metadata)
        {
            Metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
            Extend = new PropertyExtendFluent<TEntity, TProperty>(Metadata.Extend);
        }

        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        public IPropertyFluent<TEntity, TProperty> HasColumnIndex(int index)
        {
            Metadata.Column.ColumnIndex = index;
            return this;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        public IPropertyFluent<TEntity, TProperty> HasColumnWidth(int width)
        {
            Metadata.Column.Style.SetWidth(width);
            return this;
        }

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        public IPropertyFluent<TEntity, TProperty> HasColumnTitle(string title)
        {
            Metadata.Column.Title = title;
            return this;
        }

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        public IPropertyFluent<TEntity, TProperty> HasColumnFormatter(string formatter)
        {
            if(!string.IsNullOrWhiteSpace(formatter))
                Metadata.Column.Formatter = formatter;
            return this;
        }

        /// <summary>
        /// 忽略当前列
        /// </summary>
        /// <param name="ignored">是否忽略</param>
        public IPropertyFluent<TEntity, TProperty> Ignored(bool ignored = true)
        {
            Metadata.IsIgnored = ignored;
            return this;
        }

        /// <summary>
        /// 设置列输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyFluent<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc)
        {
            InternalCache.ColumnInputFormatterFuncCache.AddOrUpdate(Metadata.PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyFluent<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            InternalCache.OutputFormatterFuncCache.AddOrUpdate(Metadata.PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyFluent<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc)
        {
            InternalCache.InputFormatterFuncCache.AddOrUpdate(Metadata.PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置属性扩展
        /// </summary>
        /// <param name="action">操作</param>
        public IPropertyFluent<TEntity, TProperty> HasPropertyExtend(Action<IPropertyExtendFluent<TEntity, TProperty>> action)
        {
            action.Invoke(Extend as IPropertyExtendFluent<TEntity, TProperty>);
            return this;
        }

        /// <summary>
        /// 获取属性扩展
        /// </summary>
        public IPropertyExtendFluent<TEntity, TProperty> PropertyExtend() => Extend as IPropertyExtendFluent<TEntity, TProperty>;
    }
}
