using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Helpers;
using Bing.Utils.IO;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// Excel 导出处理
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public class ExcelExportProcess<TEntity> : IExcelExportProcess where TEntity : class, new()
    {
        /// <summary>
        /// Excel导出服务
        /// </summary>
        private readonly IExcelExportService _excelExportService;

        /// <summary>
        /// 导出选项
        /// </summary>
        private readonly IExportOptions<TEntity> _options;

        /// <summary>
        /// 导出数据函数
        /// </summary>
        private readonly GetExportDataEventAsync<TEntity> _func;

        /// <summary>
        /// 查询条件
        /// </summary>
        private readonly object _condition;

        /// <summary>
        /// 初始化一个<see cref="ExcelExportProcess{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="excelExportService">Excel导出服务</param>
        /// <param name="options">导出选项配置</param>
        public ExcelExportProcess(IExcelExportService excelExportService, IExportOptions<TEntity> options) : this(excelExportService, options, null, null) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelExportProcess{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="excelExportService">Excel导出服务</param>
        /// <param name="options">导出选项配置</param>
        /// <param name="func">导出函数事件</param>
        /// <param name="condition">查询条件</param>
        public ExcelExportProcess(IExcelExportService excelExportService, IExportOptions<TEntity> options, GetExportDataEventAsync<TEntity> func, object condition)
        {
            _excelExportService = excelExportService ?? throw new ArgumentNullException(nameof(excelExportService));
            _options = options ?? throw new ArgumentNullException(nameof(options));
            _func = func;
            _condition = condition;
        }

        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名</param>
        public async Task<ExportResult> RunAsync(string baseDir, string fileName) =>
            await RunAsync(baseDir, fileName, string.Empty);

        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名</param>
        /// <param name="exportFields">导出字段。以","分割</param>
        public async Task<ExportResult> RunAsync(string baseDir, string fileName, string exportFields)
        {
            if (_options.Data == null && _func == null)
                throw new OfficeException("无可导出的数据");
            if (_options.Data == null && _func != null)
                _options.Data = (await _func(_condition, _options.QueryCount)).ToList();
            try
            {
                return await CreateFileAsync(baseDir, fileName, exportFields);
            }
            catch (Exception e)
            {
                throw new OfficeException("生成Excel文件失败", e);
            }
        }

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名。不包含扩展名</param>
        /// <param name="exportFields">导出字段</param>
        private async Task<ExportResult> CreateFileAsync(string baseDir, string fileName, string exportFields)
        {
            baseDir = baseDir.TrimEnd('\\').TrimEnd('/');
            baseDir += "/";
            string absFilePath = Common.GetPhysicalPath(baseDir);
            DirectoryHelper.CreateIfNotExists(absFilePath);

            var newFileName = $"{fileName}_{DateTime.Now.ToString("yyyyMMddHHmmss_fff")}.xlsx";

            var filePath = Path.Combine(baseDir, newFileName);
            absFilePath = Path.Combine(absFilePath, newFileName);
            var bytes = await _excelExportService.ExportAsync(_options);
            File.WriteAllBytes(absFilePath, bytes);
            return new ExportResult()
            {
                FileName = fileName,
                Extension = ".xlsx",
                AbsFilePath = absFilePath,
                FilePath = filePath
            };
        }
    }
}
