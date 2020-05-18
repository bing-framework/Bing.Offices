using System.Collections.Generic;
using Bing.Offices.Imports;

namespace Bing.Offices.Filters
{
    /// <summary>
    /// 导入列头筛选器<br />
    /// 可以自行处理列头设置、值映射等
    /// </summary>
    public interface IImportHeaderFilter
    {
        /// <summary>
        /// 处理列头
        /// </summary>
        /// <param name="importHeaderInfos">导入列头信息集合</param>
        List<ImportHeaderInfo> Filter(List<ImportHeaderInfo> importHeaderInfos);
    }
}
