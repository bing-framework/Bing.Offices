using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Settings;

namespace Bing.Offices.Configurations
{
    /// <summary>
    /// 工作表配置
    /// </summary>
    internal class SheetConfiguration : ISheetConfiguration
    {
        /// <summary>
        /// 工作表设置
        /// </summary>
        internal SheetSetting SheetSetting { get; }

        /// <summary>
        /// 初始化一个<see cref="SheetConfiguration"/>类型的实例
        /// </summary>
        public SheetConfiguration() : this(null) { }

        /// <summary>
        /// 初始化一个<see cref="SheetConfiguration"/>类型的实例
        /// </summary>
        /// <param name="sheetSetting">工作表设置</param>
        public SheetConfiguration(SheetSetting sheetSetting)
        {
            SheetSetting = sheetSetting ?? new SheetSetting();
        }

        /// <summary>
        /// 设置索引
        /// </summary>
        /// <param name="index">索引</param>
        public ISheetConfiguration HasIndex(int index)
        {
            if (index >= 0)
            {
                SheetSetting.Index = index;
            }

            return this;
        }

        /// <summary>
        /// 设置工作表名称
        /// </summary>
        /// <param name="sheetName">名称</param>
        public ISheetConfiguration HasTitle(string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                SheetSetting.Name = sheetName;
            }

            return this;
        }

        /// <summary>
        /// 设置起始行索引
        /// </summary>
        /// <param name="index">索引</param>
        public ISheetConfiguration HasStartRowIndex(int index)
        {
            SheetSetting.StartRowIndex = index;
            return this;
        }
    }
}
