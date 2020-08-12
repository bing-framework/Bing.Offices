namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 工作表元数据
    /// </summary>
    internal sealed class SheetMetadata
    {
        /// <summary>
        /// 起始行索引
        /// </summary>
        private int _startRowIndex = 1;

        /// <summary>
        /// 工作表名称
        /// </summary>
        private string _sheetName = "Sheet0";

        /// <summary>
        /// 工作表索引
        /// </summary>
        private int _sheetIndex ;

        /// <summary>
        /// 索引
        /// </summary>
        public int Index
        {
            get => _sheetIndex;
            set
            {
                if (value >= 0)
                    _sheetIndex = value;
            }
        }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name
        {
            get => _sheetName;
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                    _sheetName = value;
            }
        }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 是否拥有标题
        /// </summary>
        public bool IsHasTitle { get; set; }

        /// <summary>
        /// 起始行索引
        /// </summary>
        public int StartRowIndex
        {
            get => _startRowIndex;
            set
            {
                if (value >= 0)
                    _startRowIndex = value;
            }
        }

        /// <summary>
        /// 标题行索引
        /// </summary>
        public int HeaderRowIndex => StartRowIndex - 1;

        /// <summary>
        /// 启用自动列宽
        /// </summary>
        public bool AutoColumnWidthEnabled { get; set; }

        /// <summary>
        /// 标题样式
        /// </summary>
        public ColumnStyleMetadata TitleStyle { get; set; }

        /// <summary>
        /// 表头样式
        /// </summary>
        public ColumnStyleMetadata HeaderStyle { get; set; }

        /// <summary>
        /// 正文样式
        /// </summary>
        public ColumnStyleMetadata BodyStyle { get; set; }

        /// <summary>
        /// 页脚样式
        /// </summary>
        public ColumnStyleMetadata FooterStyle { get; set; }
    }
}
