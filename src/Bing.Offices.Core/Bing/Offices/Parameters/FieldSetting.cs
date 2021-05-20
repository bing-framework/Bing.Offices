namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 字段设置
    /// </summary>
    public class FieldSetting
    {
        /// <summary>
        /// 字段名称
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// 初始化一个<see cref="FieldSetting"/>类型的实例
        /// </summary>
        public FieldSetting() { }

        /// <summary>
        /// 初始化一个<see cref="FieldSetting"/>类型的实例
        /// </summary>
        /// <param name="fieldName">字段名称</param>
        /// <param name="displayName">显示名称</param>
        public FieldSetting(string fieldName, string displayName) { }
    }
}
