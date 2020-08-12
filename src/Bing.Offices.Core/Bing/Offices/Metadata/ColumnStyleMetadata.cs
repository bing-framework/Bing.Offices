namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 数据列样式元数据
    /// </summary>
    internal sealed class ColumnStyleMetadata
    {
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 是否自动宽度
        /// </summary>
        public bool IsAutoWidth { get; set; }

        /// <summary>
        /// 设置宽度
        /// </summary>
        /// <param name="width">宽度</param>
        public ColumnStyleMetadata SetWidth(int width)
        {
            if (width > 0)
            {
                Width = Width;
                IsAutoWidth = false;
            }
            else if (width == -1)
            {
                Width = -1;
                IsAutoWidth = true;
            }
            return this;
        }

        /// <summary>
        /// 设置高度
        /// </summary>
        /// <param name="height">高度</param>
        public ColumnStyleMetadata SetHeight(int height)
        {
            if (height > 0)
                Height = height;
            return this;
        }
    }
}
