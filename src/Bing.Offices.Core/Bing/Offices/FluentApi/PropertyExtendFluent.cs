using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Metadata;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 属性扩展设置
    /// </summary>
    internal class PropertyExtendFluent : IPropertyExtendFluent
    {
        /// <summary>
        /// 属性扩展元数据
        /// </summary>
        public PropertyExtendMetadata Metadata { get; protected set; }
    }

    /// <summary>
    /// 属性扩展设置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal sealed class PropertyExtendFluent<TEntity, TProperty> : PropertyExtendFluent, IPropertyExtendFluent<TEntity, TProperty>
    {
        /// <summary>
        /// 初始化一个<see cref="PropertyExtendFluent"/>类型的实例
        /// </summary>
        /// <param name="metadata">属性扩展元数据</param>
        public PropertyExtendFluent(PropertyExtendMetadata metadata) => Metadata = metadata??throw new ArgumentNullException(nameof(metadata));

        /// <summary>
        /// 设置保留小数位数
        /// </summary>
        /// <param name="scale">保留小数位数</param>
        public IPropertyExtendFluent<TEntity, TProperty> HasDecimalScale(byte scale = 2)
        {
            Metadata.DecimalScale = scale;
            return this;
        }

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="defaultValue">默认值</param>
        public IPropertyExtendFluent<TEntity, TProperty> HasDefaultValue(object defaultValue)
        {
            Metadata.DefaultValue = defaultValue;
            return this;
        }

        /// <summary>
        /// 设置动态列
        /// </summary>
        /// <param name="columns">动态列</param>
        public IPropertyExtendFluent<TEntity, TProperty> HasDynamicColumn(IEnumerable<string> columns)
        {
            if (Metadata.DynamicColumns == null)
                Metadata.DynamicColumns = new List<string>();
            var list = columns.ToList();
            Metadata.DynamicColumns = list.Any() ? list : new List<string>();
            return this;
        }
    }
}
