using System;
using Bing.Offices.Metadata;
using Bing.Offices.Settings;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    internal sealed class SheetFluent : ISheetFluent
    {
        /// <summary>
        /// 冻结设置
        /// </summary>
        public FreezeSetting FreezeSetting { get; set; }

        /// <summary>
        /// 过滤器设置
        /// </summary>
        public FilterSetting FilterSetting { get; set; }

        /// <summary>
        /// 工作表元数据
        /// </summary>
        public SheetMetadata Metadata { get; set; }

        /// <summary>
        /// 初始化一个<see cref="SheetFluent"/>类型的实例
        /// </summary>
        /// <param name="metadata">工作表元数据</param>
        public SheetFluent(SheetMetadata metadata) =>
            Metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));

        /// <summary>
        /// 设置冻结窗格。
        /// 创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都会被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        public ISheetFluent HasFreezePane(int colSplit, int rowSplit)
        {
            FreezeSetting = new FreezeSetting(colSplit, rowSplit);
            return this;
        }

        /// <summary>
        /// 设置冻结窗格。
        /// 创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都会被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        /// <param name="leftMostColumn">可见列</param>
        /// <param name="topRow">可见行</param>
        public ISheetFluent HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            FreezeSetting = new FreezeSetting(colSplit, rowSplit, leftMostColumn, topRow);
            return this;
        }

        /// <summary>
        /// 设置筛选器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        public ISheetFluent HasFilter(int firstColumn) => HasFilter(firstColumn, null);

        /// <summary>
        /// 设置筛选器
        /// </summary>
        /// <param name="fistColumn">首列索引</param>
        /// <param name="lastColumn">尾列索引</param>
        public ISheetFluent HasFilter(int fistColumn, int? lastColumn)
        {
            FilterSetting = new FilterSetting(fistColumn, lastColumn);
            return this;
        }
    }
}
