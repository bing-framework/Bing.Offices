using System;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义属性元数据
    /// </summary>
    public interface IPropertyMetadata
    {
        /// <summary>
        /// 属性名称
        /// </summary>
        string PropertyName { get; set; }

        /// <summary>
        /// 属性类型
        /// </summary>
        Type PropertyType { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        bool IsIgnored { get; set; }

        /// <summary>
        /// 是否自动清空空格
        /// </summary>
        bool IsAutoTrim { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        string Description { get; set; }

        /// <summary>
        /// 数据列元数据
        /// </summary>
        IColumnMetadata Column { get; set; }
    }
}
