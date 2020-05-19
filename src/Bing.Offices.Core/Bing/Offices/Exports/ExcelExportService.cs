using System.Threading.Tasks;
using Bing.Offices.Decorators;
using Bing.Offices.Factories;
using Bing.Extensions;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// Excel导出服务
    /// </summary>
    public class ExcelExportService : IExcelExportService
    {
        /// <summary>
        /// Excel导出提供程序
        /// </summary>
        private readonly IExcelExportProvider _excelExportProvider;

        /// <summary>
        /// 初始化一个<see cref="ExcelExportService"/>类型的实例
        /// </summary>
        /// <param name="excelExportProvider">Excel导出提供程序</param>
        public ExcelExportService(IExcelExportProvider excelExportProvider) => _excelExportProvider = excelExportProvider;

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导出选项配置</param>
        public Task<byte[]> ExportAsync<T>(IExportOptions<T> options) where T : class, new()
        {
            var workbookBytes = GetExportProvider(options).Export(options);
            workbookBytes = HandleDecorate(workbookBytes, options);
            return Task.FromResult(workbookBytes);
        }

        /// <summary>
        /// 获取导出提供程序
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导出选项配置</param>
        private IExcelExportProvider GetExportProvider<T>(IExportOptions<T> options) where T:class,new()
        {
            if (options.CustomExportProvider != null)
                return options.CustomExportProvider;
            return _excelExportProvider;
        }

        /// <summary>
        /// 处理装饰
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="workbookBytes">工作簿字节数组</param>
        /// <param name="options">导出选项配置</param>
        private byte[] HandleDecorate<T>(byte[] workbookBytes, IExportOptions<T> options) where T : class, new()
        {
            var provider = GetExportProvider(options);
            var context = new DecoratorContext()
            {
                TypeDecoratorInfo = TypeDecoratorInfoFactory.CreateInstance(typeof(T))
            };
            var decorators = DecoratorFactory.CreateInstances<T>();
            decorators?.ForEach(x => { workbookBytes = x.Handler(workbookBytes, options, context, provider); });
            return workbookBytes;
        }
    }
}
