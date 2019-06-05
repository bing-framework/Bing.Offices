namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 冻结设置
    /// </summary>
    public interface IFreezeSetting
    {
        /// <summary>
        /// 冻结单元格的列号
        /// </summary>
        int ColSplit { get; }

        /// <summary>
        /// 冻结单元格的行号
        /// </summary>
        int RowSplit { get; }

        /// <summary>
        /// 顶列索引
        /// </summary>
        int LeftMostColumn { get; }

        /// <summary>
        /// 顶行索引
        /// </summary>
        int TopRow { get; }
    }
}
