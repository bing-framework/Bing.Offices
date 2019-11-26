using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Bing.Offices.Core.Models;

namespace Bing.Offices.Abstractions
{
    /// <summary>
    /// 导出器
    /// </summary>
    public interface IExporter
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="fileName">文件名</param>
        /// <param name="data">数据</param>
        Task<TemplateFileInfo> ExportAsync<T>(string fileName, ICollection<T> data) where T : class;

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="data">数据</param>
        Task<byte[]> ExportAsBytesAsync<T>(ICollection<T> data) where T : class;

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="fileName">文件名</param>
        /// <param name="data">数据</param>
        Task<TemplateFileInfo> ExportAsync<T>(string fileName, DataTable data) where T : class;

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="data">数据</param>
        Task<byte[]> ExportAsBytesAsync<T>(DataTable data) where T : class;

        /// <summary>
        /// 导出表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <param name="globalStyle">全局样式</param>
        /// <param name="styles">样式</param>
        Task<byte[]> ExportHeaderAsBytesAsync(string[] items, string sheetName, ExcelHeadStyle globalStyle = null,
            List<ExcelHeadStyle> styles = null);

        /// <summary>
        /// 导出表头
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="type">类型</param>
        Task<byte[]> ExportHeaderAsBytesAsync<T>(T type) where T : class;
    }
}
