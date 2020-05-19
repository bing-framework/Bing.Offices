namespace Bing.Offices.Metadata.Excels.Internal
{
    /// <summary>
    /// 索引范围
    /// </summary>
    internal class IndexRange
    {
        #region 属性

        /// <summary>
        /// 当前索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 结束索引
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// 是否已结束
        /// </summary>
        public bool IsEnd => Index >= EndIndex;

        #endregion

        #region 构造函数

        /// <summary>
        /// 初始化一个<see cref="IndexRange"/>类型的实例
        /// </summary>
        /// <param name="index">当前索引</param>
        /// <param name="endIndex">结束索引</param>
        public IndexRange(int index, int endIndex)
        {
            Index = index;
            EndIndex = endIndex;
        }

        #endregion

        #region GetIndex(获取索引)

        /// <summary>
        /// 获取索引
        /// </summary>
        /// <param name="span">跨度</param>
        public int GetIndex(int span = 1)
        {
            var result = Index;
            Index += span;
            return result;
        }

        #endregion

        #region Contains(判断是否包含该索引)

        /// <summary>
        /// 判断是否包含该索引
        /// </summary>
        /// <param name="index">索引</param>
        public bool Contains(int index) => index >= Index && index <= EndIndex;

        #endregion

        #region Split(分割索引范围)

        /// <summary>
        /// 分割索引范围
        /// </summary>
        /// <param name="index">索引</param>
        /// <param name="span">跨度</param>
        public IndexRange Split(int index, int span)
        {
            if (index == Index)
            {
                Index = index + span;
                return null;
            }

            var result = new IndexRange(index + span, EndIndex);
            EndIndex = index - 1;
            return result;
        }

        #endregion
    }
}
