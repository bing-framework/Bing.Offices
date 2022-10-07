using Bing.Extensions;
using Bing.Offices.Attributes;
using Bing.Offices.Decorators;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Factories;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Resolvers;
using Bing.Offices.Settings;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Collections;

namespace Bing.Offices.Npoi.Exports
{
    /// <summary>
    /// Excel导出提供程序
    /// </summary>
    public class ExcelExportProvider : IExcelExportProvider
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导出选项配置</param>
        public byte[] Export<T>(IExportOptions<T> options) where T : class, new()
        {
            var workbook = ExcelHelper.PrepareWorkbook(options.ExportFormat);
            var sheet = workbook.CreateSheet(options.SheetName);
            var headerDict = ExportMappingFactory.CreateInstance(typeof(T));
            BuildDynamicHeader(headerDict, options.DynamicColumns);
            HandleHeader(sheet, options, headerDict);
            if (options.Data != null && options.Data.Count > 0)
                HandleBody(sheet, options.DataRowStartIndex, options.Data, headerDict);
            CustomHeaderRow(sheet, options);
            return workbook?.SaveToBuffer();
        }

        /// <summary>
        /// 构建动态表头
        /// </summary>
        /// <param name="headerDict">表头映射字典</param>
        /// <param name="dynamicColumns">动态列</param>
        private void BuildDynamicHeader(IDictionary<string, PropertySetting> headerDict, IList<string> dynamicColumns)
        {
            if (!dynamicColumns.Any())
                return;
            var propertySetting = headerDict.Values.FirstOrDefault(x => x.IsDynamicColumn);
            propertySetting?.SetDynamicColumn(dynamicColumns);
        }

        /// <summary>
        /// 自定义表头行
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="options"></param>
        private void CustomHeaderRow<T>(NPOI.SS.UserModel.ISheet sheet, IExportOptions<T> options) where T : class, new()
        {
            if (options.HeaderRow == null || !options.HeaderRow.Any())
                return;
            CellRangeAddress(sheet, options);

            options.HeaderRow.ForEach(x =>
            {
                var hRow = sheet.GetRow(x.RowIndex) ?? sheet.CreateRow(x.RowIndex);
                x.Cells.OrderBy(m => m.ColumnIndex).ForEach(m =>
                {
                    var col = hRow.GetCell(m.ColumnIndex) ?? hRow.CreateCell(m.ColumnIndex);
                    col.SetCellValue(m.Value.ToString());
                });
            });
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <param name="options"></param>
        private void CellRangeAddress<T>(NPOI.SS.UserModel.ISheet sheet, IExportOptions<T> options) where T : class, new()
        {
            //if (options.HeaderRow.All(x => x.Cells.All(m => m.ColumnSpan == 1)))
            //    return;
            //var rows = options.HeaderRow.Where(x => x.Cells.Any(m => m.ColumnSpan > 1)).ToList();
            var rows = options.HeaderRow.Where(x => x.Cells.Any(t => t.RowSpan > 1 || t.ColumnSpan > 1)).ToList();
            if (rows.Any())
            {
                rows.OrderBy(x => x.RowIndex).ForEach(row =>
                {
                    row.Cells.OrderBy(x => x.ColumnIndex).ForEach(x =>
                    {
                        if (x.RowSpan <= 1 && x.ColumnSpan <= 1)
                            return;
                        var region =
                            new CellRangeAddress(x.RowIndex, x.EndRowIndex, x.ColumnIndex, x.EndColumnIndex);
                        sheet.AddMergedRegion(region);
                    });
                });
            }
        }

        /// <summary>
        /// 处理表头
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="headerDict">表头映射字典</param>
        private void HandleHeader<T>(NPOI.SS.UserModel.ISheet sheet, IExportOptions<T> options, IDictionary<string, PropertySetting> headerDict) where T : class, new()
        {
            var row = sheet.CreateRow(options.HeaderRowIndex);
            var columnIndex = 0;
            foreach (var kvp in headerDict)
            {
                if (kvp.Value.Ignored)
                    continue;
                if (kvp.Value.IsDynamicColumn)
                {
                    foreach (var column in kvp.Value.DynamicColumns)
                    {
                        row.CreateCell(columnIndex).SetCellValue(column);
                        columnIndex++;
                    }
                    continue;
                }
                row.CreateCell(columnIndex).SetCellValue(kvp.Value.Title);
                columnIndex++;
            }
        }

        /// <summary>
        /// 处理正文
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="dataRowStartIndex">数据行起始索引</param>
        /// <param name="data">数据集</param>
        /// <param name="headerDict">表头映射字典</param>
        private void HandleBody<T>(NPOI.SS.UserModel.ISheet sheet, int dataRowStartIndex, IList<T> data,
            IDictionary<string, PropertySetting> headerDict) where T : class, new()
        {
            if (data.Count <= 0)
                return;
            for (var i = 0; i < data.Count; i++)
            {
                var columnIndex = 0;
                var row = sheet.CreateRow(dataRowStartIndex + i);
                var dto = data[i];
                foreach (var kvp in headerDict)
                {
                    if (kvp.Value.Ignored)
                        continue;
                    if (kvp.Value.IsDynamicColumn)
                    {
                        var dictionary = dto.GetExtendDictionary(kvp.Key);
                        foreach (var column in kvp.Value.DynamicColumns)
                        {
                            if (column.IsEmpty())
                                continue;
                            row.CreateCell(columnIndex).SetCellValue(dictionary[column]?.ToString());
                            columnIndex++;
                        }
                        continue;
                    }
                    var cell = row.CreateCell(columnIndex);
                    if (!string.IsNullOrWhiteSpace(kvp.Value.Formatter))
                        cell.SetCellValue(dto.GetStringValue(kvp.Key, kvp.Value.Formatter));
                    else if (kvp.Value.MappingValues.Any() && kvp.Value.MappingValues.Count > 0)
                    {
                        var value = dto.GetValue(kvp.Key);
                        if (kvp.Value.MappingValues.ContainsValue(value))
                        {
                            var mapValue = kvp.Value.MappingValues.FirstOrDefault(f => f.Value.Equals(value));
                            cell.SetCellValue(mapValue.Key);
                        }
                    }
                    else
                        cell.SetValue(dto.GetValue(kvp.Key), kvp.Value.DecimalScale);
                    columnIndex++;
                }
            }
        }

        /// <summary>
        /// 处理表头
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿流</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="context">装饰器上下文</param>
        public byte[] HandleHeader<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context) where T : class, new()
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            var attribute =
                context.TypeDecoratorInfo?.TypeDecorators?.SingleOrDefault(x => x.GetType() == typeof(HeaderAttribute))
                    as HeaderAttribute;
            if (attribute == null)
                return workbookBytes;
            var workbook = workbookBytes.ToWorkbook();
            var headerRow = workbook?.GetSheet(options.SheetName)?.GetRow(options.HeaderRowIndex);
            if (headerRow == null)
                return workbookBytes;
            var style = workbook.CreateCellStyle();
            var font = workbook.CreateFont();
            font.FontName = attribute.FontName;
            font.Color = ColorResolver.Resolve(attribute.Color);
            font.FontHeightInPoints = attribute.FontSize;
            font.IsBold = attribute.Bold;
            style.SetFont(font);
            for (var i = 0; i < headerRow.PhysicalNumberOfCells; i++)
                headerRow.GetCell(i).CellStyle = style;
            return workbook.SaveToBuffer();
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿流</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="context">装饰器上下文</param>
        public byte[] MergeCells<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context) where T : class, new()
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            var propertyDecoratorInfos = context.TypeDecoratorInfo.PropertyDecoratorInfos;
            var workbook = workbookBytes.ToWorkbook();
            var sheet = workbook.GetSheet(options.SheetName);
            foreach (var item in propertyDecoratorInfos)
            {
                if (item.Decorators.SingleOrDefault(x => x.GetType() == typeof(MergeColumnsAttribute)) != null)
                {
                    MergeCells(sheet, item.ColumnIndex, options);
                }
            }

            return workbook.SaveToBuffer();
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">工作表</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="options">导出选项配置</param>
        private void MergeCells<T>(NPOI.SS.UserModel.ISheet sheet, int columnIndex, IExportOptions<T> options)
            where T : class, new()
        {
            string currentCellValue;
            var startRowIndex = options.DataRowStartIndex;
            NPOI.SS.Util.CellRangeAddress mergeRangeAddress;
            var startRow = sheet.GetRow(startRowIndex);
            if (startRow == null)
                return;
            var startCell = startRow.GetCell(columnIndex);
            if (startCell == null)
                return;
            string startCellValue = startCell.StringCellValue;
            if (string.IsNullOrWhiteSpace(startCellValue))
                return;

            for (var rowIndex = options.DataRowStartIndex; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
            {
                var cell = sheet.GetRow(rowIndex)?.GetCell(columnIndex);
                currentCellValue = cell == null ? string.Empty : cell.StringCellValue;
                if (currentCellValue.Trim() != startCellValue.Trim())
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex - 1, columnIndex, columnIndex);
                    if (mergeRangeAddress.NumberOfCells > 1)
                    {
                        sheet.AddMergedRegion(mergeRangeAddress);
                        startRow.GetCell(columnIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    }

                    startRowIndex = rowIndex;
                    startCellValue = currentCellValue;
                }

                if (rowIndex == sheet.PhysicalNumberOfRows - 1 && startRowIndex != rowIndex)
                {
                    mergeRangeAddress = new CellRangeAddress(startRowIndex, rowIndex, columnIndex, columnIndex);
                    sheet.AddMergedRegion(mergeRangeAddress);
                    startRow.GetCell(columnIndex).CellStyle.VerticalAlignment = VerticalAlignment.Center;
                }
            }
        }

        /// <summary>
        /// 自动换行
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿流</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="context">装饰器上下文</param>
        public byte[] WarpText<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context) where T : class, new()
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));
            var attribute = context.TypeDecoratorInfo.GetDecoratorAttribute<WrapTextAttribute>();
            if (attribute == null)
                return workbookBytes;
            var workbook = workbookBytes.ToWorkbook();
            var sheet = workbook.GetSheet(options.SheetName);
            if (sheet.PhysicalNumberOfRows <= 0)
                return workbookBytes;

            for (var i = 0; i < sheet.PhysicalNumberOfRows; i++)
            {
                var row = sheet.GetRow(i);
                for (var colIndex = 0; colIndex < row.PhysicalNumberOfCells; colIndex++)
                    row.GetCell(colIndex).CellStyle.WrapText = true;
            }
            return workbook.SaveToBuffer();
        }
    }
}