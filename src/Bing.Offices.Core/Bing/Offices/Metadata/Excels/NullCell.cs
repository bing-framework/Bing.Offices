namespace Bing.Offices.Metadata.Excels
{
    /// <summary>
    /// 空单元格
    /// </summary>
    public class NullCell : Cell, ICell
    {
        /// <summary>
        /// 初始化一个<see cref="NullCell"/>类型的实例
        /// </summary>
        public NullCell() : base("")
        {
        }

        /// <summary>
        /// 是否为空单元格
        /// </summary>
        public override bool IsNull() => true;
    }
}
