namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 工作表设置
    /// </summary>
    public interface ISheetSetting
    {
        /// <summary>
        /// 工作表索引
        /// </summary>
        int Index { get; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 起始行索引
        /// </summary>
        int StartRowIndex { get; }

        /// <summary>
        /// 标题行索引
        /// </summary>
        int HeaderRowIndex { get; }
    }
}
