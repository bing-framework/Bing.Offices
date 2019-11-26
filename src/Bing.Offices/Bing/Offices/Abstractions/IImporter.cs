using System.Threading.Tasks;
using Bing.Offices.Core.Models;

namespace Bing.Offices.Abstractions
{
    /// <summary>
    /// 导入器
    /// </summary>
    public interface IImporter
    {
        /// <summary>
        /// 生成导入模板
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="fileName">文件名</param>
        Task<TemplateFileInfo> GenerateTemplateAsync<T>(string fileName) where T : class, new();

        /// <summary>
        /// 生成导入模板
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        Task<byte[]> GenerateTemplateAsBytesAsync<T>() where T : class, new();

        /// <summary>
        /// 导入
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="filePath">文件路径</param>
        Task<ImportResult<T>> ImportAsync<T>(string filePath) where T : class, new();
    }
}
