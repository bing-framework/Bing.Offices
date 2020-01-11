using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Metadata.Excels;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Metadata.Excels;
using Bing.Utils.Extensions;
using Bing.Utils.Helpers;
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
        /// 哈希表动态单元格
        /// </summary>
        private static readonly Hashtable TableDynamicCell = Hashtable.Synchronized(new Hashtable(1024));

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
        public IWorkbook Convert<TTemplate>(string fileUrl, int sheetIndex = 0, int headerRowIndex = 0, int dataRowStartIndex = 1, bool multiSheet = false, int maxColumnLength = 100) where TTemplate : class, new()
        {
            var workbook = new NpoiWorkbook();
            var innerWorkbook = GetWorkbook(fileUrl);
            if (multiSheet == false)
            {
                BuildSheet<TTemplate>(workbook, innerWorkbook, sheetIndex, headerRowIndex, dataRowStartIndex,
                    maxColumnLength);
                return workbook;
            }

            for (var i = 0; i < innerWorkbook.NumberOfSheets; i++)
            {
                BuildSheet<TTemplate>(workbook, innerWorkbook, i, headerRowIndex, dataRowStartIndex, maxColumnLength);
            }

            return workbook;
        }

        /// <summary>
        /// 构建工作表
        /// </summary>
        /// <typeparam name="TTemplate">导入模板类型</typeparam>
        /// <param name="workbook">工作簿</param>
        /// <param name="innerWorkbook">内部工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="headerRowIndex">表头行索引</param>
        /// <param name="dataRowStartIndex">数据行起始索引</param>
        /// <param name="maxColumnLength">最大列长度</param>
        private void BuildSheet<TTemplate>(Bing.Offices.Abstractions.Metadata.Excels.IWorkbook workbook, NPOI.SS.UserModel.IWorkbook innerWorkbook, int sheetIndex, int headerRowIndex, int dataRowStartIndex, int maxColumnLength)
        {
            var innerSheet = GetSheet(innerWorkbook, sheetIndex);
            if (innerSheet.GetRow(0).PhysicalNumberOfCells > maxColumnLength * 10)
                throw new OfficeException($"导入数据初始化过多的无效列: {maxColumnLength}");
            var sheet = workbook.CreateSheet(innerSheet.SheetName, headerRowIndex);
            HandleHeader(sheet, innerSheet, headerRowIndex);
            HandleBody<TTemplate>(sheet, innerSheet, dataRowStartIndex);
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
                cells.Add(new Cell(innerCell.GetStringValue()) { ColumnIndex = i, Name = innerCell.GetStringValue() });
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
            var hash = string.Empty;
            if (type.GetCustomAttribute<HasDynamicColumnAttribute>() != null)
                hash = $"_{Encrypt.Md5By32(header.Cells.Select(x => x.Name).Join())}";

            string columnName;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                columnName = header?.Cells?.SingleOrDefault(x => x.ColumnIndex == i)?.Name;
                if (string.IsNullOrWhiteSpace(columnName))
                    continue;
                var key = $"{type.FullName}_{hash}_{i}";

                if (Table[key] == null)
                {
                    var isDynamic = false;
                    // 优先匹配列名特性值
                    var matchProperty = props.FirstOrDefault(x =>
                        x.GetCustomAttribute<ColumnNameAttribute>()?.Name == columnName);
                    // 次之匹配属性名
                    if (matchProperty == null)
                        matchProperty = props.FirstOrDefault(x =>
                            x.Name.Equals(columnName, StringComparison.CurrentCultureIgnoreCase));
                    // 次之匹配动态列
                    if (matchProperty == null)
                    {
                        matchProperty = props.FirstOrDefault(x => x.GetCustomAttribute<DynamicColumnAttribute>() != null);
                        if (matchProperty != null)
                            isDynamic = true;
                    }
                    var propertyName = matchProperty?.Name;
                    Table[key] = propertyName;
                    TableDynamicCell[key] = isDynamic;
                }

                var value = row.GetCell(i) == null ? string.Empty : row.GetCell(i).GetStringValue();
                var cell = new Cell(value)
                {
                    ColumnIndex = i,
                    Name = columnName,
                    PropertyName = Table[key]?.ToString(),
                    IsDynamic = Bing.Utils.Helpers.Conv.ToBool(TableDynamicCell[key])
                };
                cells.Add(cell);
            }

            return cells;
        }
    }
}
