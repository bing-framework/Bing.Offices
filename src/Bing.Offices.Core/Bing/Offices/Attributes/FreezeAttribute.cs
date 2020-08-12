using System;
using Bing.Offices.Settings;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 冻结窗格 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class FreezeAttribute : Attribute
    {
        /// <summary>
        /// 冻结设置
        /// </summary>
        internal FreezeSetting FreezeSetting { get; }

        /// <summary>
        /// 垂直分割位置
        /// </summary>
        public int ColSplit => FreezeSetting.ColSplit;

        /// <summary>
        /// 水平分割位置
        /// </summary>
        public int RowSplit => FreezeSetting.RowSplit;

        /// <summary>
        /// 可见列数
        /// </summary>
        public int LeftMostColumn => FreezeSetting.LeftMostColumn;

        /// <summary>
        /// 可见行数
        /// </summary>
        public int TopRow => FreezeSetting.TopRow;

        /// <summary>
        /// 初始化一个<see cref="FreezeAttribute"/>类型的实例
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        public FreezeAttribute(int colSplit, int rowSplit) : this(colSplit, rowSplit, 0, 1) { }

        /// <summary>
        /// 初始化一个<see cref="FreezeAttribute"/>类型的实例
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        /// <param name="leftMostColumn">可见列数</param>
        /// <param name="topRow">可见行数</param>
        public FreezeAttribute(int colSplit, int rowSplit, int leftMostColumn, int topRow) => FreezeSetting = new FreezeSetting(colSplit, rowSplit, leftMostColumn, topRow);
    }
}
