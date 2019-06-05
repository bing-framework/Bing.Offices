using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices.Settings
{
    /// <summary>
    /// 工作表设置
    /// </summary>
    public sealed class SheetSetting : ISheetSetting
    {
        /// <summary>
        /// 工作表索引
        /// </summary>
        private int _index = 0;

        /// <summary>
        /// 工作表名称
        /// </summary>
        private string _name = "Sheet0";

        /// <summary>
        /// 起始行索引
        /// </summary>
        private int _startRowIndex = 1;

        /// <summary>
        /// 工作表索引
        /// </summary>
        public int Index
        {
            get => _index;
            internal set => _index = value >= 0 ? value : 0;
        }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string Name
        {
            get => _name;
            internal set => _name = string.IsNullOrWhiteSpace(value) ? value : "Sheet0";
        }

        /// <summary>
        /// 起始行索引
        /// </summary>
        public int StartRowIndex
        {
            get => _startRowIndex;
            internal set => _startRowIndex = value >= 1 ? value : 1;
        }

        /// <summary>
        /// 标题行索引
        /// </summary>
        public int HeaderRowIndex => StartRowIndex + 1;
    }
}
