using System.Collections.Generic;
using Bing.Offices.Decorators;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Abstractions
{
    /// <summary>
    /// Excel导出器提供程序
    /// </summary>
    public interface IExcelExportProvider
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据</param>
        byte[] Export<T>(List<T> data) where T : class, new();

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="exportOptions">导出选项</param>
        byte[] Export<T>(ExportOptions<T> exportOptions) where T : class, new();

        /// <summary>
        /// 装饰表头
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="exportOptions">导出选项</param>
        /// <param name="context">装饰器上下文</param>
        byte[] DecorateHeader<T>(byte[] workbookBytes, ExportOptions<T> exportOptions, DecoratorContext context)
            where T : class, new();

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="exportOptions">导出选项</param>
        /// <param name="context">装饰器上下文</param>
        byte[] MergeCell<T>(byte[] workbookBytes, ExportOptions<T> exportOptions, DecoratorContext context)
            where T : class, new();

        /// <summary>
        /// 自动换行
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="exportOptions">导出选项</param>
        /// <param name="context">装饰器上下文</param>
        byte[] WrapText<T>(byte[] workbookBytes, ExportOptions<T> exportOptions, DecoratorContext context)
            where T : class, new();
    }
}
