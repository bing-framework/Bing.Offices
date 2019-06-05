namespace Bing.Offices.Abstractions.Configurations
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    public interface ISheetConfiguration
    {
        /// <summary>
        /// 设置索引
        /// </summary>
        /// <param name="index">索引</param>
        ISheetConfiguration HasIndex(int index);

        /// <summary>
        /// 设置工作表名称
        /// </summary>
        /// <param name="sheetName">名称</param>
        ISheetConfiguration HasTitle(string sheetName);

        /// <summary>
        /// 设置起始行索引
        /// </summary>
        /// <param name="index">索引</param>
        ISheetConfiguration HasStartRowIndex(int index);
    }
}
