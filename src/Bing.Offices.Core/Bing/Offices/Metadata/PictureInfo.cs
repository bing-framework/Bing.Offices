namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 图片信息
    /// </summary>
    public class PictureInfo
    {
        /// <summary>
        /// 最小行索引
        /// </summary>
        public int MinRow { get; set; }

        /// <summary>
        /// 最大行索引
        /// </summary>
        public int MaxRow { get; set; }

        /// <summary>
        /// 最小列索引
        /// </summary>
        public int MinCol { get; set; }

        /// <summary>
        /// 最大列索引
        /// </summary>
        public int MaxCol { get; set; }

        /// <summary>
        /// 图片数据
        /// </summary>
        public byte[] PictureData { get; set; }

        /// <summary>
        /// 图片样式
        /// </summary>
        public PictureStyle PictureStyle { get; set; }

        /// <summary>
        /// 初始化一个<see cref="PictureInfo"/>类型的实例
        /// </summary>
        /// <param name="minRow">最小行索引</param>
        /// <param name="maxRow">最大行索引</param>
        /// <param name="minCol">最小列索引</param>
        /// <param name="maxCol">最大列索引</param>
        /// <param name="pictureData">图片数据</param>
        /// <param name="pictureStyle">图片样式</param>
        public PictureInfo(int minRow, int maxRow, int minCol, int maxCol, byte[] pictureData,
            PictureStyle pictureStyle)
        {
            MinRow = minRow;
            MaxRow = maxRow;
            MinCol = minCol;
            MaxCol = maxCol;
            PictureData = pictureData;
            PictureStyle = pictureStyle;
        }
    }
}
