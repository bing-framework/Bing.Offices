namespace Bing.Offices.Abstractions.Metadata.Excels
{
    /// <summary>
    /// 单元格
    /// </summary>
    public interface ICell
    {
        /// <summary>
        /// 值
        /// </summary>
        object Value { get; set; }

        /// <summary>
        /// 行
        /// </summary>
        IRow Row { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        int ColumnSpan { get; set; }

        /// <summary>
        /// 行跨度
        /// </summary>
        int RowSpan { get; }

        /// <summary>
        /// 列索引
        /// </summary>
        int ColumnIndex { get; set; }

        /// <summary>
        /// 行索引
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// 结束列索引
        /// </summary>
        int EndColumnIndex { get; }

        /// <summary>
        /// 结束行索引
        /// </summary>
        int EndRowIndex { get; }

        /// <summary>
        /// 是否需要合并单元格。true:是,false:否
        /// </summary>
        bool NeedMerge { get; }

        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 属性名称
        /// </summary>
        string PropertyName { get; set; }

        /// <summary>
        /// 是否动态单元格
        /// </summary>
        bool IsDynamic { get; set; }

        /// <summary>
        /// 是否为空单元格
        /// </summary>
        bool IsNull();
    }
}
