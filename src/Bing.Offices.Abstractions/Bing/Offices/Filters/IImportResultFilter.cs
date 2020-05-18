using Bing.Offices.Imports;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 导入结果筛选器 <br />
    /// 可以处理标注内容
    /// </summary>
    public interface IImportResultFilter
    {
        /// <summary>
        /// 处理导入结果 <br />
        /// 比如：对错误信息进行多语言转换
        /// </summary>
        /// <typeparam name="T">结果类型</typeparam>
        /// <param name="importResult">导入结果</param>
        ImportResult<T> Filter<T>(ImportResult<T> importResult) where T : class, new();
    }
}
