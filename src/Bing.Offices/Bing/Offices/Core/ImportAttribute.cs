using System;

namespace Bing.Offices.Core
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
    }
}
