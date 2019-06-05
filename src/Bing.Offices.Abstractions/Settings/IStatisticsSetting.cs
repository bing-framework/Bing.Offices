namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 统计信息设置
    /// </summary>
    public interface IStatisticsSetting
    {
        /// <summary>
        /// 统计信息名称，默认名称位置为（最后一行，第一个单元格）
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 单元格公式，如：SUM、AVERAGE等。可用于垂直统计
        /// </summary>
        string Formula { get; }

        /// <summary>
        /// 统计信息的列索引。如果<see cref="Formula"/>是SUM，而<see cref="Columns"/>是[1,3]，例如：列1和列3将是SUM第一行到最后一行。
        /// </summary>
        int[] Columns { get; }
    }
}
