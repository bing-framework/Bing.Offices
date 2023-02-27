using Bing.Offices.Metadata.Excels;
using System.Collections.Generic;

namespace Bing.Offices.Exports;

/// <summary>
/// 导出选项配置
/// </summary>
/// <typeparam name="T">实体类型</typeparam>
public interface IExportOptions<T> where T : class, new()
{
    /// <summary>
    /// 行头
    /// </summary>
    IList<IRow> HeaderRow { get; set; }

    /// <summary>
    /// 数据集合
    /// </summary>
    IList<T> Data { get; set; }

    /// <summary>
    /// 导出格式
    /// </summary>
    ExcelFormat ExportFormat { get; set; }

    /// <summary>
    /// 工作表名称
    /// </summary>
    string SheetName { get; set; }

    /// <summary>
    /// 表头行索引。默认为0
    /// </summary>
    int HeaderRowIndex { get; set; }

    /// <summary>
    /// 数据行起始索引。默认为1
    /// </summary>
    int DataRowStartIndex { get; set; }

    /// <summary>
    /// 动态列
    /// </summary>
    IList<string> DynamicColumns { get; set; }

    /// <summary>
    /// 自定义导出提供程序
    /// </summary>
    IExcelExportProvider CustomExportProvider { get; set; }

    /// <summary>
    /// 查询总数
    /// </summary>
    int QueryCount { get; set; }
}