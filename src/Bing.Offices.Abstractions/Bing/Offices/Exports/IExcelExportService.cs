using System.Threading.Tasks;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// Excel导出服务
    /// </summary>
    public interface IExcelExportService
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导出选项配置</param>
        Task<byte[]> ExportAsync<T>(IExportOptions<T> options) where T : class, new();
    }
}
