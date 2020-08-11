using Bing.Offices.Extensions;
using System;
using System.Reflection;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 属性元数据
    /// </summary>
    internal class PropertyMetadata : IPropertyMetadata
    {
        /// <summary>
        /// 属性名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 属性类型
        /// </summary>
        public Type PropertyType { get; set; }

        /// <summary>
        /// 属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnored { get; set; }

        /// <summary>
        /// 数据列元数据
        /// </summary>
        public ColumnMetadata Column { get; set; }

        /// <summary>
        /// 属性扩展元数据
        /// </summary>
        public PropertyExtendMetadata Extend { get; set; }
    }

    /// <summary>
    /// 属性元数据
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal sealed class PropertyMetadata<TEntity, TProperty> : PropertyMetadata, IPropertyMetadata<TEntity, TProperty>
    {
        /// <summary>
        /// 初始化一个<see cref="PropertyMetadata"/>类型的实例
        /// </summary>
        /// <param name="propertyInfo">属性信息</param>
        public PropertyMetadata(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;
            PropertyName = propertyInfo.Name;
            PropertyType = typeof(TProperty);
            Column = new ColumnMetadata
            {
                Title = propertyInfo.Name
            };
            Extend = new PropertyExtendMetadata<TEntity, TProperty>();
        }

        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        public IPropertyMetadata<TEntity, TProperty> HasColumnIndex(int index)
        {
            Column.ColumnIndex = index;
            return this;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        public IPropertyMetadata<TEntity, TProperty> HasColumnWidth(int width)
        {
            Column.Style.SetWidth(width);
            return this;
        }

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        public IPropertyMetadata<TEntity, TProperty> HasColumnTitle(string title)
        {
            Column.Title = title;
            return this;
        }

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        public IPropertyMetadata<TEntity, TProperty> HasColumnFormatter(string formatter)
        {
            if (string.IsNullOrWhiteSpace(formatter))
                Column.Formatter = formatter;
            return this;
        }

        /// <summary>
        /// 忽略当前列
        /// </summary>
        /// <param name="ignored">是否忽略</param>
        public IPropertyMetadata<TEntity, TProperty> Ignored(bool ignored = true)
        {
            IsIgnored = ignored;
            return this;
        }

        /// <summary>
        /// 设置保留小数位数
        /// </summary>
        /// <param name="scale">保留小数位数</param>
        public IPropertyMetadata<TEntity, TProperty> HasDecimalScale(byte scale = 2)
        {
            Extend.DecimalScale = scale;
            return this;
        }

        /// <summary>
        /// 设置列输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyMetadata<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc)
        {
            InternalCache.ColumnInputFormatterFuncCache.AddOrUpdate(PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyMetadata<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            InternalCache.OutputFormatterFuncCache.AddOrUpdate(PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyMetadata<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc)
        {
            InternalCache.InputFormatterFuncCache.AddOrUpdate(PropertyInfo, formatterFunc);
            return this;
        }

        /// <summary>
        /// 设置属性扩展
        /// </summary>
        /// <param name="action">操作</param>
        public IPropertyMetadata<TEntity, TProperty> HasPropertyExtend(Action<IPropertyExtendMetadata<TEntity, TProperty>> action)
        {
            action.Invoke(Extend as PropertyExtendMetadata<TEntity, TProperty>);
            return this;
        }

        /// <summary>
        /// 获取属性扩展
        /// </summary>
        public IPropertyExtendMetadata<TEntity, TProperty> PropertyExtend() => Extend as PropertyExtendMetadata<TEntity, TProperty>;
    }
}
