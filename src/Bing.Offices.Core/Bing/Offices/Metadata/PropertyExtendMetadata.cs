using System.Collections.Generic;
using System.Linq;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 属性扩展元数据
    /// </summary>
    internal class PropertyExtendMetadata : IPropertyExtendMetadata
    {
        /// <summary>
        /// 保留小数位数
        /// </summary>
        public byte? DecimalScale { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; set; }

        /// <summary>
        /// 动态列
        /// </summary>
        public List<string> DynamicColumns { get; set; }

    }

    /// <summary>
    /// 属性扩展元数据
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal sealed class PropertyExtendMetadata<TEntity, TProperty> : PropertyExtendMetadata, IPropertyExtendMetadata<TEntity, TProperty>
    {
        /// <summary>
        /// 是否动态列
        /// </summary>
        private bool _isDynamicColumn;

        /// <summary>
        /// 初始化一个<see cref="PropertyExtendMetadata{TEntity,TProperty}"/>类型的实例
        /// </summary>
        public PropertyExtendMetadata()
        {
            _isDynamicColumn = false;
        }

        /// <summary>
        /// 设置保留小数位数
        /// </summary>
        /// <param name="scale">保留小数位数</param>
        public IPropertyExtendMetadata<TEntity, TProperty> HasDecimalScale(byte scale = 2)
        {
            DecimalScale = scale;
            return this;
        }

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="defaultValue">默认值</param>
        public IPropertyExtendMetadata<TEntity, TProperty> HasDefaultValue(object defaultValue)
        {
            DefaultValue = defaultValue;
            return this;
        }

        /// <summary>
        /// 设置动态列
        /// </summary>
        /// <param name="columns">动态列</param>
        public IPropertyExtendMetadata<TEntity, TProperty> HasDynamicColumn(IEnumerable<string> columns)
        {
            if (DynamicColumns == null)
                DynamicColumns = new List<string>();
            var list = columns.ToList();
            if (list.Any())
            {
                DynamicColumns = list;
                _isDynamicColumn = true;
            }
            else
            {
                _isDynamicColumn = false;
            }
            return this;
        }
    }
}
