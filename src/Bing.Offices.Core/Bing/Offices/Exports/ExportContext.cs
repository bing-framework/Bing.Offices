using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Internals;
using Bing.Offices.Settings;

namespace Bing.Offices.Exports;

/// <summary>
/// 导出上下文
/// </summary>
public class ExportContext<TEntity>
{
    /// <summary>
    /// 初始化一个<see cref="ExportContext{TEntity}"/>类型的实例
    /// </summary>
    /// <param name="excelFormat"></param>
    /// <param name="sheetIndex"></param>
    /// <param name="configuration">Excel 配置</param>
    internal ExportContext(ExcelFormat excelFormat, int sheetIndex, ExcelConfiguration<TEntity> configuration)
    {
        Format = excelFormat;
        SheetIndex = sheetIndex;
        Configuration = configuration;
    }

    /// <summary>
    /// Excel 配置
    /// </summary>
    internal ExcelConfiguration<TEntity> Configuration { get; }

    /// <summary>
    /// Excel格式
    /// </summary>
    public ExcelFormat Format { get; }

    /// <summary>
    /// 工作簿索引
    /// </summary>
    public int SheetIndex { get; }

    /// <summary>
    /// Excel 设置
    /// </summary>
    public ExcelSetting ExcelSetting => Configuration.ExcelSetting;

    /// <summary>
    /// 全局工作表设置
    /// </summary>
    public GlobalSheetSetting GlobalSheetSetting => Configuration.GlobalSheetSetting;

    /// <summary>
    /// 冻结设置
    /// </summary>
    public IList<FreezeSetting> FreezeSettings => Configuration.FreezeSettings;

    /// <summary>
    /// 导出属性设置
    /// </summary>
    public IDictionary<PropertyInfo, ExportColumnPropertySetting> PropertySettings =>
        Configuration.PropertyConfigurationDictionary.ToDictionary(x => x.Key, x => x.Value.ColumnPropertySetting);

    /// <summary>
    /// 获取输出格式化函数
    /// </summary>
    /// <param name="propertyInfo">属性信息</param>
    public Delegate GetOutputFormatterFunc(PropertyInfo propertyInfo)
    {
        if (InternalCache.OutputFormatterFuncCache.TryGetValue(propertyInfo, out var formatterFunc) && formatterFunc?.Method != null)
            return formatterFunc;
        return null;
    }

    /// <summary>
    /// 获取默认的工作表设置
    /// </summary>
    public SheetSettingBase GetDefaultSheetSetting()
    {
        var currentSheetSetting = GetSheetSetting(SheetIndex);
        if (currentSheetSetting == null)
            return GlobalSheetSetting;
        return currentSheetSetting;
    }

    /// <summary>
    /// 获取指定索引的工作表设置
    /// </summary>
    /// <param name="sheetIndex">工作表索引</param>
    public SheetSetting GetSheetSetting(int sheetIndex)
    {
        if (Configuration.SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
            return sheetSetting;
        return null;
    }
}
