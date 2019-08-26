using Bing.Offices.Abstractions.Metadata.Excels;

namespace Bing.Offices.Abstractions.Imports
{
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
        IWorkbook Convert<TTemplate>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0,
            int dataRowStartIndex = 1) where TTemplate : class, new();
    }
}
