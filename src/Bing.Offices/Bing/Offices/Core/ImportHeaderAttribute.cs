using System;

namespace Bing.Offices.Core
{
    /// <summary>
    /// 导入头部 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImportHeaderAttribute : Attribute
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 批注
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; } = "jian玄冰";

        /// <summary>
        /// 自动过滤空格。默认启用
        /// </summary>
        public bool AutoTrim { get; set; } = true;

        /// <summary>
        /// 处理掉所有的空格，包括中间空格
        /// </summary>
        public bool FixAllSpace { get; set; }

        /// <summary>
        /// 列索引。如果为0则自动计算
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 是否允许重复。默认启用
        /// </summary>
        public bool IsAllowRepeat { get; set; } = true;

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }
    }
}
