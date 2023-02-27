using System.Collections.Generic;

namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 工作簿
/// </summary>
public interface IWorkbook
{
    /// <summary>
    /// 工作表列表
    /// </summary>
    IList<IWorkSheet> Sheets { get; }

    /// <summary>
    /// 工作表
    /// </summary>
    /// <param name="sheetIndex">工作表索引</param>
    IWorkSheet this[int sheetIndex] { get; }

    /// <summary>
    /// 获取工作表
    /// </summary>
    /// <param name="sheetName">工作表名称</param>
    IWorkSheet GetSheet(string sheetName);

    /// <summary>
    /// 获取工作表
    /// </summary>
    /// <param name="sheetIndex">工作表索引</param>
    IWorkSheet GetSheetAt(int sheetIndex);

    /// <summary>
    /// 创建工作表
    /// </summary>
    IWorkSheet CreateSheet();

    /// <summary>
    /// 创建工作表
    /// </summary>
    /// <param name="sheetName">工作表名称</param>
    IWorkSheet CreateSheet(string sheetName);

    /// <summary>
    /// 创建工作表
    /// </summary>
    /// <param name="sheetName">工作表名称</param>
    /// <param name="startRowIndex">起始行索引</param>
    IWorkSheet CreateSheet(string sheetName, int startRowIndex);

    /// <summary>
    /// 添加工作表
    /// </summary>
    /// <param name="sheet">工作表</param>
    void AddSheet(IWorkSheet sheet);
}