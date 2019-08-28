using Bing.Offices.Abstractions.Decorators;
using Bing.Offices.Abstractions.Exports;
using Bing.Offices.Attributes;

namespace Bing.Offices.Decorators
{
    /// <summary>
    /// 表头装饰器
    /// </summary>
    [BindDecorator(typeof(HeaderAttribute))]
    internal class HeaderDecorator : IDecorator
    {
        /// <summary>
        /// 处理
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="context">装饰器上下文</param>
        /// <param name="excelExportProvider">Excel导出提供程序</param>
        public byte[] Handler<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context,
            IExcelExportProvider excelExportProvider) where T : class, new() =>
            excelExportProvider.HandleHeader(workbookBytes, options, context);
    }
}
