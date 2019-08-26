using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Attributes;
using Bing.Offices.Metadata.Excels;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Metadata.Excels;
using NPOI.SS.UserModel;
using IWorkbook = Bing.Offices.Abstractions.Metadata.Excels.IWorkbook;
using IRow = Bing.Offices.Abstractions.Metadata.Excels.IRow;
using ICell = Bing.Offices.Abstractions.Metadata.Excels.ICell;

namespace Bing.Offices.Npoi.Imports
{
    /// <summary>
    /// Excel导入提供程序
    /// </summary>
    public class ExcelImportProvider : IExcelImportProvider
    {
        /// <summary>
        /// 哈希表
        /// </summary>
        private static readonly Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

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
            var workbook = new NpoiWorkbook();
            var innerWorkbook = GetWorkbook(fileUrl);
            for (var i = 0; i < innerWorkbook.NumberOfSheets - 1; i++)
            {
                var innerSheet = GetSheet(innerWorkbook, i);
                var sheet = workbook.CreateSheet(innerSheet.SheetName);
                HandleHeader(sheet, innerSheet, headerRowIndex);
                HandleBody<TTemplate>(sheet, innerSheet, dataRowStartIndex);
            }
            return workbook;
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

        /// <summary>
        /// 处理表头
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="innerSheet">NPOI工作表</param>
        /// <param name="headerRowIndex">标题行索引</param>
        private void HandleHeader(IWorkSheet sheet, ISheet innerSheet, int headerRowIndex)
        {
            var innerRow = innerSheet.GetRow(headerRowIndex);
            var cells = new List<ICell>();
            for (var i = 0; i < innerRow.PhysicalNumberOfCells; i++)
            {
                var innerCell = innerRow.GetCell(i);
                cells.Add(new Cell(innerCell.GetStringValue()) {ColumnIndex = i, Name = innerCell.GetStringValue()});
            }
            sheet.AddHeadRow(cells.ToArray());
        }

        /// <summary>
        /// 处理正文
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="innerSheet">NPOI工作表</param>
        /// <param name="dataRowStartIndex">数据行起始索引</param>
        private void HandleBody<TTemplate>(IWorkSheet sheet, ISheet innerSheet, int dataRowStartIndex)
        {
            var header = sheet.GetHeader().LastOrDefault();
            for (var i = dataRowStartIndex; i < innerSheet.PhysicalNumberOfRows; i++)
            {
                var innerRow = innerSheet.GetRow(i);
                if (innerRow == null || innerRow.Cells.All(x => string.IsNullOrWhiteSpace(x?.GetStringValue())))
                    continue;
                sheet.AddBodyRow(Convert<TTemplate>(innerRow, header));
            }
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <typeparam name="TTemplate">模板类型</typeparam>
        /// <param name="row">NPOI单元行</param>
        /// <param name="header">表头行</param>
        private IList<ICell> Convert<TTemplate>(NPOI.SS.UserModel.IRow row, IRow header)
        {
            var type = typeof(TTemplate);
            var props = type.GetProperties().ToList();
            var cells = new List<ICell>();

            string columnName;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                columnName = header?.Cells?.SingleOrDefault(x => x.ColumnIndex == i)?.Name;
                if (string.IsNullOrWhiteSpace(columnName))
                    continue;
                var key = $"{type.FullName}_{i}";
                if (Table[key] == null)
                {
                    // 优先匹配列名特性值
                    var matchProperty = props.FirstOrDefault(x =>
                        x.GetCustomAttribute<ColumnNameAttribute>()?.Name == columnName);
                    // 次之匹配属性名
                    if (matchProperty == null)
                        matchProperty = props.FirstOrDefault(x =>
                            x.Name.Equals(columnName, StringComparison.CurrentCultureIgnoreCase));
                    var propertyName = matchProperty?.Name;
                    Table[key] = propertyName;
                }

                var value = row.GetCell(i) == null ? string.Empty : row.GetCell(i).GetStringValue();
                var cell = new Cell(value)
                {
                    ColumnIndex = i,
                    Name = columnName,
                    PropertyName = Table[key]?.ToString(),
                };
                cells.Add(cell);
            }

            return cells;
        }

    }
}
