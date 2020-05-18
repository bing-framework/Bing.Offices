namespace Bing.Offices.Imports
{
    /// <summary>
    /// 导入头部
    /// </summary>
    public interface IImportHeader
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 批注
        /// </summary>
        string Description { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        string Author { get; set; }

        /// <summary>
        /// 自动过滤空格。默认: 启用
        /// </summary>
        bool AutoTrim { get; set; }

        /// <summary>
        /// 处理掉所有的空格，包括中间空格
        /// </summary>
        bool FixAllSpace { get; set; }

        /// <summary>
        /// 列索引。如果为0则自动计算
        /// </summary>
        int ColumnIndex { get; set; }

        /// <summary>
        /// 是否允许重复。默认: 允许
        /// </summary>
        bool IsAllowRepeat { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        bool IsIgnore { get; set; }
    }
}
