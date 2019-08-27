using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Abstractions.Decorators;
using Bing.Offices.Abstractions.Exports;
using Bing.Offices.Attributes;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Factories;
using Bing.Offices.Npoi.Resolvers;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
            var workbook = CreateWorkbook(options.ExportType);
            var sheet = workbook.CreateSheet(options.SheetName);
            var headerDict = ExportMappingFactory.CreateInstance(typeof(T));
            HandleHeader(sheet, options.HeaderRowIndex, headerDict);
            if (options.Data != null && options.Data.Count > 0)
                HandleBody(sheet, options.DataRowStartIndex, options.Data, headerDict);
            return workbook?.SaveToBuffer();
        }

        /// <summary>
        /// 创建工作簿
        /// </summary>
        /// <param name="type">导出类型</param>
        private NPOI.SS.UserModel.IWorkbook CreateWorkbook(ExportType type)
        {
            switch (type)
            {
                case ExportType.Xls:
                    return new HSSFWorkbook();
                case ExportType.Xlsx:
                    return new XSSFWorkbook();
                default:
                    return new XSSFWorkbook();
            }
        }

        /// <summary>
        /// 处理表头
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="headerRowIndex">表头行索引</param>
        /// <param name="headerDict">表头映射字典</param>
        private void HandleHeader(ISheet sheet, int headerRowIndex, IDictionary<string, string> headerDict)
        {
            var row = sheet.CreateRow(headerRowIndex);
            var columnIndex = 0;
            foreach (var kvp in headerDict)
            {
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
        private void HandleBody<T>(ISheet sheet, int dataRowStartIndex, IList<T> data,
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
            throw new System.NotImplementedException();
        }

        
    }
}
