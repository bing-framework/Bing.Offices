using Bing.Offices.Abstractions;
using Bing.Offices.Attributes.Decorators;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Decorators.Internal
{
    /// <summary>
    /// 自动换行装饰器
    /// </summary>
    [DecoratorBind(typeof(WrapTextAttribute))]
    internal class WrapTextDecorator : IDecorator
    {
        /// <summary>
        /// 装饰
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="exportOptions">导出选项</param>
        /// <param name="context">装饰器上下文</param>
        /// <param name="excelExportProvider">Excel导出器提供程序</param>
        public byte[] Decorate<T>(byte[] workbookBytes, ExportOptions<T> exportOptions, DecoratorContext context,
            IExcelExportProvider excelExportProvider) where T : class, new()
        {
            return excelExportProvider.WrapText(workbookBytes, exportOptions, context);
        }
    }
}
