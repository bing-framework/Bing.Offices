namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 导入参数
    /// </summary>
    public class ImportParam : IImportParam
    {
        /// <summary>
        /// 数据源名称
        /// </summary>
        public string DataSourceName { get; set; }

        /// <summary>
        /// 数据填充方向
        /// </summary>
        public DataDirection DataDirection { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ImportParam"/>类型的实例
        /// </summary>
        /// <param name="dataSourceName">数据源名称</param>
        public ImportParam(string dataSourceName) : this(dataSourceName, DataDirection.None)
        {
        }

        /// <summary>
        /// 初始化一个<see cref="ImportParam"/>类型的实例
        /// </summary>
        /// <param name="dataSourceName">数据源名称</param>
        /// <param name="dataDirection">数据填充方向</param>
        public ImportParam(string dataSourceName, DataDirection dataDirection)
        {
            DataSourceName = dataSourceName;
            DataDirection = dataDirection;
        }
    }
}
