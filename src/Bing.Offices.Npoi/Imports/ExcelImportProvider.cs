using System;
using System.IO;
using Bing.Offices.Abstractions.Imports;
using NPOI.SS.UserModel;
using IWorkbook = Bing.Offices.Abstractions.Metadata.Excels.IWorkbook;

namespace Bing.Offices.Npoi.Imports
{
    /// <summary>
    /// Excel导入提供程序
    /// </summary>
    public class ExcelImportProvider : IExcelImportProvider
    {
        /// <summary>
        /// 转换
        /// </summary>
        /// <typeparam name="TTemplate">导入模板类型</typeparam>
        /// <param name="fileUrl">文件地址</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="headerRowIndex">标题行索引</param>
        /// <param name="dataRowStartIndex">数据行起始索引</param>
        public IWorkbook Convert<TTemplate>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1) where TTemplate : class, new()
        {
            var workbook = GetWorkbook(fileUrl);

            throw new NotImplementedException();
        }

        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="fileUrl">文件地址</param>
        private NPOI.SS.UserModel.IWorkbook GetWorkbook(string fileUrl)
        {
            if (string.IsNullOrWhiteSpace(fileUrl))
                throw new ArgumentNullException(nameof(fileUrl));
            if (!File.Exists(fileUrl))
                throw new FileNotFoundException("找不到文件", fileUrl);
            var ext = Path.GetExtension(fileUrl).ToLower().Trim();
            if (ext != ".xls" && ext != ".xlsx")
                throw new NotSupportedException("仅支持后缀名为.xls或者.xlsx的文件");
            return WorkbookFactory.Create(fileUrl);
        }

        /// <summary>
        /// 获取工作表
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        private ISheet GetSheet(NPOI.SS.UserModel.IWorkbook workbook, int sheetIndex = 0) => workbook.GetSheetAt(sheetIndex);
    }
}
