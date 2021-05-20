using System.Collections.Generic;
using System.Threading.Tasks;
using Bing.Offices.Parameters;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// 模板导出
    /// </summary>
    public interface ITemplateExport
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="templateExportParam">模板导出参数</param>
        Task<byte[]> ExportAsync(string templatePath, ITemplateExportParam templateExportParam);

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="templatePath">模板路径</param>
        /// <param name="templateExportParams">模板导出参数列表</param>
        Task<byte[]> ExportAsync(string templatePath, List<ITemplateExportParam> templateExportParams);

        Task<byte[]> ExportAsync()
    }
}
