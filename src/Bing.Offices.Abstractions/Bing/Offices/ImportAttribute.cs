using System;
using Bing.Offices.Filters;

namespace Bing.Offices
{
    /// <summary>
    /// 导入 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ImportAttribute : Attribute
    {
        /// <summary>
        /// 表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;

        /// <summary>
        /// 最大允许导入的数量
        /// </summary>
        public int MaxCount { get; set; } = 50000;

        /// <summary>
        /// 导入结果筛选器 <br />
        /// 必须实现【<see cref="IImportResultFilter"/>】
        /// </summary>
        public Type ImportResultFilter { get; set; }

        /// <summary>
        /// 导入列头筛选器 <br />
        /// 必须实现【<see cref="IImportHeaderFilter"/>】
        /// </summary>
        public Type ImportHeaderFilter { get; set; }
    }
}
