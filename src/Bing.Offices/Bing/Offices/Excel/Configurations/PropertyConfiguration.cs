using System;
using Bing.Offices.Excel.Settings;

namespace Bing.Offices.Excel.Configurations
{
    /// <summary>
    /// 属性配置
    /// </summary>
    internal class PropertyConfiguration : IPropertyConfiguration
    {
        /// <summary>
        /// 属性设置
        /// </summary>
        internal PropertySetting PropertySetting { get; }

        /// <summary>
        /// 初始化一个<see cref="PropertyConfiguration"/>类型的实例
        /// </summary>
        /// <param name="propertySetting">属性设置</param>
        public PropertyConfiguration(PropertySetting propertySetting) => PropertySetting = propertySetting;
    }

    /// <summary>
    /// 属性配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration, IPropertyConfiguration<TEntity, TProperty>
    {
        /// <summary>
        /// 初始化一个<see cref="PropertyConfiguration{TEntity,TProperty}"/>类型的实例
        /// </summary>
        public PropertyConfiguration() : this(new PropertySetting<TEntity, TProperty>()) { }

        /// <summary>
        /// 初始化一个<see cref="PropertyConfiguration{TEntity,TProperty}"/>类型的实例
        /// </summary>
        /// <param name="propertySetting">属性设置</param>
        public PropertyConfiguration(PropertySetting<TEntity, TProperty> propertySetting) : base(propertySetting) => PropertySetting = propertySetting;

        /// <summary>
        /// 属性设置
        /// </summary>
        internal new PropertySetting<TEntity, TProperty> PropertySetting { get; }

        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="index">索引</param>
        public IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index)
        {
            if (index >= 0)
                PropertySetting.ColumnIndex = index;
            return this;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="width">宽度</param>
        public IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width)
        {
            PropertySetting.ColumnWidth = width;
            return this;
        }

        /// <summary>
        /// 设置列标题
        /// </summary>
        /// <param name="title">标题</param>
        public IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title)
        {
            if (!string.IsNullOrWhiteSpace(title))
                PropertySetting.ColumnTitle = title;
            return this;
        }

        /// <summary>
        /// 设置列格式化
        /// </summary>
        /// <param name="formatter">格式化</param>
        public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter)
        {
            if (!string.IsNullOrWhiteSpace(formatter))
                PropertySetting.ColumnFormatter = formatter;
            return this;
        }

        /// <summary>
        /// 设置忽略该属性
        /// </summary>
        public IPropertyConfiguration<TEntity, TProperty> Ignored()
        {
            PropertySetting.IsIgnored = true;
            return this;
        }

        /// <summary>
        /// 设置输出格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
        {
            if (formatterFunc != null)
                PropertySetting.OutputFormatterFunc = formatterFunc;
            return this;
        }

        /// <summary>
        /// 设置输入格式化
        /// </summary>
        /// <param name="formatterFunc">格式化函数</param>
        public IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc)
        {
            if (formatterFunc != null)
                PropertySetting.InputFormatterFunc = formatterFunc;
            return this;
        }
    }
}
