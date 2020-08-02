using System;
using System.Collections.Generic;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义类元数据
    /// </summary>
    public interface IClassMetadata
    {
        /// <summary>
        /// 类类型
        /// </summary>
        Type ClassType { get; set; }

        /// <summary>
        /// 类模式
        /// </summary>
        ClassMode Mode { get; set; }

        /// <summary>
        /// 属性元数据列表
        /// </summary>
        IList<IPropertyMetadata> Properties { get; set; }

        /// <summary>
        /// 工作表元数据
        /// </summary>
        ISheetMetadata Sheet { get; set; }
    }
}
