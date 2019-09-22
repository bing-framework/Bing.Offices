using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Abstractions.Exports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 工作簿<see cref="IWorkbook"/> 扩展
    /// </summary>
    public static class WorkbookExtensions
    {
        #region SaveToBuffer(将工作簿转换成字节数组)

        /// <summary>
        /// 将工作簿转换成字节数组
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static byte[] SaveToBuffer(this IWorkbook workbook)
        {
            using (var ms = new MemoryStream())
            {
                workbook.Write(ms);
                return ms.ToArray();
            }
        }

        #endregion

        #region ToWorkbook(转换为工作簿)

        /// <summary>
        /// 转换为工作簿
        /// </summary>
        /// <param name="workbookBytes">工作簿字节数组</param>
        public static IWorkbook ToWorkbook(this byte[] workbookBytes)
        {
            using (var stream = new MemoryStream(workbookBytes))
                return WorkbookFactory.Create(stream);
        }

        #endregion

        #region GetExcelFormat(获取Excel格式类型)

        /// <summary>
        /// 获取Excel格式类型
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static ExportFormat GetExcelFormat(this IWorkbook workbook)
        {
            switch (workbook)
            {
                case HSSFWorkbook _:
                    return ExportFormat.Xlsx;
                case XSSFWorkbook _:
                    return ExportFormat.Xls;
            }
            throw new NotImplementedException($"未知Excel格式类型");
        }

        #endregion

        #region GetSheets(获取工作表集合)

        /// <summary>
        /// 获取工作表集合
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static IEnumerable<ISheet> GetSheets(this IWorkbook workbook)
        {
            var sheets = new List<ISheet>();
            for (var i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if(sheet!=null&&!workbook.IsSheetHidden(i))
                    sheets.Add(sheet);
            }
            return sheets;
        }

        #endregion

    }
}
