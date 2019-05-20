using System.Collections.Generic;
using Bing.Offices.Models.Excels;

namespace Bing.Offices.Abstractions
{
    /// <summary>
    /// Excel导出器提供程序
    /// </summary>
    public interface IExcelExportProvider
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="data">数据</param>
        byte[] Export<T>(List<T> data) where T : class, new();

        /// <summary>
        /// 导出
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="exportOption">导出选项</param>
        byte[] Export<T>(ExportOptions<T> exportOption) where T : class, new();
    }
}
