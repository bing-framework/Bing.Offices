using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Imports;

/// <summary>
/// Excel导入提供程序
/// </summary>
public interface IExcelImportProvider
{
    /// <summary>
    /// 转换
    /// </summary>
    /// <typeparam name="TTemplate">导入模板类型</typeparam>
    /// <param name="fileUrl">文件地址</param>
    /// <param name="sheetIndex">工作表索引</param>
    /// <param name="headerRowIndex">标题行索引</param>
    /// <param name="dataRowStartIndex">数据行起始索引</param>
    /// <param name="multiSheet">是否支持多工作表模式</param>
    /// <param name="maxColumnLength">最大列长度</param>
    /// <param name="enabledEmptyLine">启用空行模式。启用时，行内遇到空行将抛出异常错误信息</param>
    IWorkbook Convert<TTemplate>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0,
        int dataRowStartIndex = 1, bool multiSheet = false, int maxColumnLength = 100, bool enabledEmptyLine = false) where TTemplate : class, new();

    /// <summary>
    /// 转换
    /// </summary>
    /// <typeparam name="TTemplate">导入模板类型</typeparam>
    /// <param name="options">导入选项配置</param>
    IWorkbook Convert<TTemplate>(IImportOptions options) where TTemplate : class, new();
}