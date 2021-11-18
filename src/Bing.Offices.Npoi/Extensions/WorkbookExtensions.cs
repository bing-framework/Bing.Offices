using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Exports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作簿(<see cref="NPOI.SS.UserModel.IWorkbook"/>) 扩展
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
            using var ms = new MemoryStream();
            workbook.Write(ms);
            return ms.ToArray();
        }

        #endregion

        #region ToWorkbook(转换为工作簿)

        /// <summary>
        /// 转换为工作簿
        /// </summary>
        /// <param name="workbookBytes">工作簿字节数组</param>
        public static IWorkbook ToWorkbook(this byte[] workbookBytes)
        {
            using var stream = new MemoryStream(workbookBytes);
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

        #region SetAllSheetAutoCompute(设置所有工作表自动计算)

        /// <summary>
        /// 设置所有工作表自动计算
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static void SetAllSheetAutoCompute(this IWorkbook workbook)
        {
            if (workbook.NumberOfSheets <= 0)
                return;
            for (var i = 0; i < workbook.NumberOfSheets; i++)
                workbook.GetSheetAt(i).ForceFormulaRecalculation = true;// 让公式自动计算
        }

        #endregion

        #region AddSheet(添加工作表)

        /// <summary>
        /// 添加工作表
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="name">工作表名称</param>
        /// <param name="heads">表头</param>
        public static ISheet AddSheet(this IWorkbook workbook,string name,List<string> heads)
        {
            var sheet = workbook.CreateSheet(name);
            var style = workbook.DefaultHeadStyle();
            var row = sheet.CreateRow(0);
            row.Height = 20 * 20;
            heads.ForEach(item => row.Value(heads.IndexOf(item), item, style));
            return sheet;
        }

        #endregion

        #region DefaultHeadStyle(默认表头样式)

        /// <summary>
        /// 默认表头样式
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static ICellStyle DefaultHeadStyle(this IWorkbook workbook)
        {
            var style = workbook.CreateCellStyle();
            var font = workbook.CreateFont();
            font.IsBold = true;// 加粗

            style.FillForegroundColor = 13;// 13为黄色
            style.FillPattern = FillPattern.SolidForeground;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.SetFont(font.DefaultFont());
            return style;
        }

        #endregion

        #region DefaultBodyStyle(默认正文样式)

        /// <summary>
        /// 默认正文样式
        /// </summary>
        /// <param name="workbook">工作簿</param>
        public static ICellStyle DefaultBodyStyle(this IWorkbook workbook)
        {
            var style = workbook.CreateCellStyle();
            var font = workbook.CreateFont();
            style.Alignment = HorizontalAlignment.Center;
            style.VerticalAlignment = VerticalAlignment.Center;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;
            style.SetFont(font.DefaultFont());
            return style;
        }

        #endregion
    }
}
