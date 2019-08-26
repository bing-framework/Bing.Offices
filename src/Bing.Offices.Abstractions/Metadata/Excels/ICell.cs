namespace Bing.Offices.Abstractions.Metadata.Excels
{
    /// <summary>
    /// 单元格
    /// </summary>
    public interface ICell
    {
        /// <summary>
        /// 单元格索引
        /// </summary>
        int Index { get; set; }

        /// <summary>
        /// 单元行索引
        /// </summary>
        int RowIndex { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 属性名
        /// </summary>
        string PropertyName { get; set; }

        /// <summary>
        /// 单元格值
        /// </summary>
        object Value { get; set; }

        /// <summary>
        /// 单元行
        /// </summary>
        IRow Row { get; set; }

        /// <summary>
        /// 设置单元格值
        /// </summary>
        /// <param name="value">值</param>
        void SetValue(object value);
    }
}
