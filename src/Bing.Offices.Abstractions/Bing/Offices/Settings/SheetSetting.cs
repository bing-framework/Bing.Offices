namespace Bing.Offices.Settings
{
    /// <summary>
    /// 工作表设置
    /// </summary>
    public sealed class SheetSetting
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        private string _sheetName = "Sheet0";

        /// <summary>
        /// 起始行索引
        /// </summary>
        private int _startRowIndex = 1;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName
        {
            get => _sheetName;
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                    _sheetName = value;
            }
        }

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
        /// 结束行索引
        /// </summary>
        public int? EndRowIndex { get; set; }

        /// <summary>
        /// 启用自动列宽
        /// </summary>
        public bool AutoColumnWidthEnabled { get; set; }
    }
}
