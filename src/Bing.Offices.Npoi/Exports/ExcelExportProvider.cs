using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Abstractions.Decorators;
using Bing.Offices.Abstractions.Exports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Attributes;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Factories;
using Bing.Offices.Npoi.Resolvers;
using Bing.Utils.Extensions;

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
            var workbook = CreateWorkbook(options.ExportFormat);
            var sheet = workbook.CreateSheet(options.SheetName);
            var headerDict = ExportMappingFactory.CreateInstance(typeof(T));
            BuildDynamicHeader(headerDict, options.DynamicColumns);
            HandleHeader(sheet, options.HeaderRowIndex, headerDict);
            if (options.Data != null && options.Data.Count > 0)
                HandleBody(sheet, options.DataRowStartIndex, options.Data, headerDict);
            return workbook?.SaveToBuffer();
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="format">导出格式</param>
        private NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExportFormat format)
        {
            switch (format)
            {
                case ExportFormat.Xls:
                    return new NPOI.HSSF.UserModel.HSSFWorkbook();
                case ExportFormat.Xlsx:
                    return new NPOI.XSSF.UserModel.XSSFWorkbook();
            }
            throw new NotImplementedException();
        }

        /// <summary>
        /// 构建动态表头
        /// </summary>
        /// <param name="headerDict">表头映射字典</param>
        /// <param name="dynamicColumns">动态列</param>
        private void BuildDynamicHeader(IDictionary<string, string> headerDict, IList<string> dynamicColumns)
        {
            if (!dynamicColumns.Any())
                return;
            var dynamicKey = headerDict.Keys.FirstOrDefault(x => x.StartsWith("Dynamic:"));
            if (string.IsNullOrWhiteSpace(dynamicKey))
                return;
            headerDict[dynamicKey] = dynamicColumns.Join(',');
        }

        /// <summary>
        /// 处理表头
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="headerRowIndex">表头行索引</param>
        /// <param name="headerDict">表头映射字典</param>
        private void HandleHeader(NPOI.SS.UserModel.ISheet sheet, int headerRowIndex, IDictionary<string, string> headerDict)
        {
            var row = sheet.CreateRow(headerRowIndex);
            var columnIndex = 0;
            foreach (var kvp in headerDict)
            {
                if (kvp.Key.StartsWith("Dynamic:"))
                {
                    var dynamicColumns = kvp.Value.Split(',');
                    foreach (var column in dynamicColumns)
                    {
                        row.CreateCell(columnIndex).SetCellValue(column);
                        columnIndex++;
                    }
                    continue;
                }
                row.CreateCell(columnIndex).SetCellValue(kvp.Value);
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
            IDictionary<string, string> headerDict) where T : class, new()
        {
            if (data.Count <= 0)
                return;
            for (int i = 0; i < data.Count; i++)
            {
                var columnIndex = 0;
                var row = sheet.CreateRow(dataRowStartIndex + i);
                var dto = data[i];
                foreach (var kvp in headerDict)
                {
                    if (kvp.Key.StartsWith("Dynamic:"))
                    {
                        var extendPropertyName = kvp.Key.Split(':')[1];
                        var dynamicColumns = kvp.Value.Split(',');
                        var dictionary = dto.GetExtendDictionary(extendPropertyName);
                        foreach (var column in dynamicColumns)
                        {
                            if (column.IsEmpty())
                            {
                                continue;
                            }
                            row.CreateCell(columnIndex).SetCellValue(dictionary[column]?.ToString());
                            columnIndex++;
                        }
                        continue;
                    }
                    row.CreateCell(columnIndex).SetCellValue(dto.GetStringValue(kvp.Key));
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
            if (attribute.Bold)
                font.Boldweight = short.MaxValue;
            style.SetFont(font);
            for (int i = 0; i < headerRow.PhysicalNumberOfCells; i++)
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
            throw new System.NotImplementedException();
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
