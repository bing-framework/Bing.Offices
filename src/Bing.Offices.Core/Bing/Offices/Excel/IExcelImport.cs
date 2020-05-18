using System.Collections.Generic;
using System.Threading.Tasks;
using Bing.Offices.Imports;

namespace Bing.Offices.Excel
{
    /// <summary>
    /// Excel 导入
    /// </summary>
    public interface IExcelImport : IImport
    {
        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="businessErrorDataList">错误数据</param>
        /// <param name="message">成功：错误数据返回路径，失败：返回错误原因</param>
        bool OutputBusinessErrorData<T>(string filePath, List<DataRowErrorInfo> businessErrorDataList,
            out string message) where T : class, new();

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath) where T : class, new();

        /// <summary>
        /// 导入多个相同类型的Sheet数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <typeparam name="TSheet">实体类型</typeparam>
        /// <param name="filePath">文件路径</param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型TSheet</returns>
        Task<Dictionary<string, ImportResult<TSheet>>> ImportMultipleSheet<T, TSheet>(string filePath)
            where T : class, new() where TSheet : class, new();
    }
}
