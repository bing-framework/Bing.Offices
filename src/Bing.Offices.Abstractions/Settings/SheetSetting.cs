namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 工作表设置
    /// </summary>
    public sealed class SheetSetting
    {
        /// <summary>
        /// 工作表索引
        /// </summary>
        private int _sheetIndex = 0;

        /// <summary>
        /// 工作表名称
        /// </summary>
        private string _sheetName = "Sheet0";

        /// <summary>
        /// 起始行索引
        /// </summary>
        private int _startRowIndex = 1;
        
        /// <summary>
        /// 工作表索引
        /// </summary>
        public int SheetIndex
        {
            get => _sheetIndex;
            set => _sheetIndex = value >= 0 ? value : 0;
        }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName
        {
            get => _sheetName;
            set => _sheetName = string.IsNullOrWhiteSpace(value) ? value : "Sheet0";
        }

        /// <summary>
        /// 起始行索引
        /// </summary>
        public int StartRowIndex
        {
            get => _startRowIndex;
            set => _startRowIndex = value >= 1 ? value : 1;
        }

        /// <summary>
        /// 标题行索引
        /// </summary>
        public int HeaderRowIndex => StartRowIndex + 1;
    }
}
