using System.Collections.Generic;

namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 直接导出参数
    /// </summary>
    public class DirectExportParam : ExportParamBase, IExportParam
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 字段设置列表
        /// </summary>
        public IReadOnlyList<FieldSetting> FieldSettings { get; private set; }

        /// <summary>
        /// 初始化一个<see cref="DirectExportParam"/>类型的实例
        /// </summary>
        /// <param name="dataSource">数据源</param>
        public DirectExportParam(object dataSource)
        {
            DataSource = dataSource;
        }

        /// <summary>
        /// 初始化一个<see cref="DirectExportParam"/>类型的实例
        /// </summary>
        /// <param name="dataSource">数据源</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="fieldSettings">字段设置列表</param>
        public DirectExportParam(object dataSource, string sheetName, List<FieldSetting> fieldSettings)
        {
            DataSource = dataSource;
            SheetName = sheetName;
            FieldSettings = fieldSettings;
        }

        /// <summary>
        /// 初始化字段设置
        /// </summary>
        /// <param name="fieldSettings">字段设置列表</param>
        public void InitFields(List<FieldSetting> fieldSettings)
        {
            FieldSettings = fieldSettings;
        }
    }
}
