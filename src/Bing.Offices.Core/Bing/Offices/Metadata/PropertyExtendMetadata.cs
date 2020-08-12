using System.Collections.Generic;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 属性扩展元数据
    /// </summary>
    internal sealed class PropertyExtendMetadata
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
}
