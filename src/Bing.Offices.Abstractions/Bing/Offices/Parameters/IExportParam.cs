using System;

namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 导出参数
    /// </summary>
    public interface IExportParam
    {
        /// <summary>
        /// 对象类型
        /// </summary>
        Type ObjectType { get; }

        /// <summary>
        /// 记录数
        /// </summary>
        int RecordCount { get; }

        /// <summary>
        /// 数据源类型
        /// </summary>
        DataSourceType DataSourceType { get; }
    }
}
