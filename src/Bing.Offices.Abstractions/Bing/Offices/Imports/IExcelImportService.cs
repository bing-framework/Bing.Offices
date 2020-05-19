using System.Threading.Tasks;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Imports
{
    /// <summary>
    /// Excel导入服务
    /// </summary>
    public interface IExcelImportService
    {
        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="options">导入选项配置</param>
        Task<IWorkbook> ImportAsync<T>(IImportOptions options) where T : class, new();
    }
}
