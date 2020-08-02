namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义数据列元数据
    /// </summary>
    public interface IColumnMetadata
    {
        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        string Formatter { get; set; }

        /// <summary>
        /// 合并单元格
        /// </summary>
        bool HasMergeCells { get; set; }

        /// <summary>
        /// 自动索引值
        /// </summary>
        bool HasAutoIndex { get; set; }

        /// <summary>
        /// 单元格索引
        /// </summary>
        int ColumnIndex { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        int RowIndex { get; set; }

        /// <summary>
        /// 范围
        /// </summary>
        CellRange Range { get; set; }

        /// <summary>
        /// 样式
        /// </summary>
        IColumnStyleMetadata Style { get; set; }
    }
}
