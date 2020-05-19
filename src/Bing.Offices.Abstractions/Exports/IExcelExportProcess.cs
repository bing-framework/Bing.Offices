using System.Collections.Generic;
using System.Threading.Tasks;

namespace Bing.Offices.Abstractions.Exports
{
    /// <summary>
    /// Excel 导出处理
    /// </summary>
    public interface IExcelExportProcess
    {
        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名</param>
        Task<ExportResult> RunAsync(string baseDir, string fileName);

        /// <summary>
        /// 执行方法
        /// </summary>
        /// <param name="baseDir">基础路径</param>
        /// <param name="fileName">文件名</param>
        /// <param name="exportFields">导出字段。以","分割</param>
        Task<ExportResult> RunAsync(string baseDir, string fileName, string exportFields);
    }

    /// <summary>
    /// 获取导出数据事件
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="condition">查询条件</param>
    /// <param name="count">数量</param>
    public delegate Task<IEnumerable<T>> GetExportDataEventAsync<T>(object condition, int count);
}
