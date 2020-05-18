using System.Threading.Tasks;
using Bing.Offices.Exports;

namespace Bing.Offices.Imports
{
    /// <summary>
    /// 导入
    /// </summary>
    public interface IImport
    {
        /// <summary>
        /// 生成导入模板
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="fileName">文件名称</param>
        Task<ExportFileInfo> GenerateTemplateAsync<T>(string fileName) where T : class, new();

        /// <summary>
        /// 生成导入模板
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        Task<byte[]> GenerateTemplateBytesAsync<T>() where T : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="labelingFilePath">标注文件路径</param>
        Task<ImportResult<T>> ImportAsync<T>(string filePath, string labelingFilePath = null) where T : class, new();
    }
}
