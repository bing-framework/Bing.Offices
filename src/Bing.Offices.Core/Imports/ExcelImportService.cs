using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Bing.Offices.Abstractions.Imports;
using Bing.Offices.Abstractions.Metadata.Excels;
using Bing.Offices.Factories;
using Bing.Offices.Filters;

namespace Bing.Offices.Imports
{
    /// <summary>
    /// Excel导入服务
    /// </summary>
    public class ExcelImportService : IExcelImportService
    {
        /// <summary>
        /// Excel导入提供程序
        /// </summary>
        private readonly IExcelImportProvider _excelImportProvider;

        /// <summary>
        /// 初始一个<see cref="ExcelImportService"/>类型的实例
        /// </summary>
        /// <param name="excelImportProvider">Excel导入提供程序</param>
        public ExcelImportService(IExcelImportProvider excelImportProvider)
        {
            _excelImportProvider = excelImportProvider;
        }

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导入选项配置</param>
        public Task<IWorkbook> ImportAsync<T>(IImportOptions options) where T : class, new()
        {
            var workbook = GetWorkbook<T>(options);
            if (options.MappingDictionary != null)
                MappingHeaderDictionary(workbook, options.MappingDictionary);
            var andFilter = new AndFilter()
            {
                Filters = FilterFactory.CreateInstances<T>()
            };
            var context = new FilterContext
            {
                TypeFilterInfo = TypeFilterInfoFactory.CreateInstance(typeof(T))
            };
            andFilter.Filter(workbook, context, options);
            return Task.FromResult(workbook);
        }

        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导入选项配置</param>
        private IWorkbook GetWorkbook<T>(IImportOptions options) where T : class, new()
        {
            return options.CustomImportProvider == null
                ? _excelImportProvider.Convert<T>(options.FileUrl, options.SheetIndex, options.HeaderRowIndex,
                    options.DataRowIndex)
                : options.CustomImportProvider.Convert<T>(options.FileUrl, options.SheetIndex, options.HeaderRowIndex,
                    options.DataRowIndex);
        }

        /// <summary>
        /// 映射表头字典数据
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="headerDictionary">表头映射字典</param>
        private void MappingHeaderDictionary(IWorkbook workbook, IDictionary<string, string> headerDictionary)
        {
            foreach (var kvp in headerDictionary)
            {
                foreach (var sheet in workbook.Sheets)
                {
                    foreach (var row in sheet.GetBody())
                    {
                        foreach (var cell in row.Cells)
                        {
                            if (cell.Name.Equals(kvp.Value, StringComparison.CurrentCultureIgnoreCase))
                                cell.PropertyName = kvp.Key;
                        }
                    }
                }
            }
        }
    }
}
