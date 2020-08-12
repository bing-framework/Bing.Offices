using System;
using System.Reflection;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 属性元数据
    /// </summary>
    internal sealed class PropertyMetadata
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
        public ColumnMetadata Column { get; set; } = new ColumnMetadata();

        /// <summary>
        /// 属性扩展元数据
        /// </summary>
        public PropertyExtendMetadata Extend { get; set; }

        /// <summary>
        /// 初始化一个<see cref="PropertyMetadata"/>类型的实例
        /// </summary>
        public PropertyMetadata() => Extend = new PropertyExtendMetadata();

        /// <summary>
        /// 初始化一个<see cref="PropertyMetadata"/>类型的实例
        /// </summary>
        /// <param name="propertyInfo">属性信息</param>
        public PropertyMetadata(PropertyInfo propertyInfo)
        {
            PropertyInfo = propertyInfo;
            PropertyName = propertyInfo.Name;
            PropertyType = propertyInfo.PropertyType;
            Column = new ColumnMetadata
            {
                Title = propertyInfo.Name
            };
            Extend = new PropertyExtendMetadata();
        }
    }
}
