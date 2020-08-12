using System.Collections.Generic;

namespace Bing.Offices.Metadata.Excels
{
    /// <summary>
    /// 工作簿
    /// </summary>
    public interface IWorkbook
    {
        /// <summary>
        /// 工作表数量
        /// </summary>
        int SheetCount { get; }

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        ISheet GetSheet(int sheetIndex);

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        ISheet CreateSheet(string sheetName);

        /// <summary>
        /// 转换为字节数组
        /// </summary>
        byte[] ToBytes();

        /// <summary>
        /// 工作表列表
        /// </summary>
        IList<ISheet> Sheets { get; }

        /// <summary>
        /// 工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        ISheet this[int sheetIndex] { get; }

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        ISheet GetSheet(string sheetName);

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        ISheet GetSheetAt(int sheetIndex);

        /// <summary>
        /// 创建工作表
        /// </summary>
        ISheet CreateSheet();

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        ISheet CreateSheet(string sheetName, int startRowIndex);

        /// <summary>
        /// 添加工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        void AddSheet(ISheet sheet);
    }
}
