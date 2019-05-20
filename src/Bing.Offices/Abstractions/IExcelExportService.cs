using System.Threading.Tasks;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Abstractions
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
        /// <param name="exportOption">导出选项</param>
        Task<byte[]> ExportAsync<T>(ExportOptions<T> exportOption) where T : class, new();
    }
}
