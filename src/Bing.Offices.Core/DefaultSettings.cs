using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices
{
    /// <summary>
    /// 默认设置
    /// </summary>
    public static class DefaultSettings
    {
        /// <summary>
        /// 默认Excel文档属性设置
        /// </summary>
        public static ExcelSetting DefaultExcelSetting { get; } = new ExcelSetting();
    }
}
