namespace Bing.Offices.Settings;

/// <summary>
/// 全局工作表设置
/// </summary>
public sealed class GlobalSheetSetting : SheetSettingBase
{
    /// <summary>
    /// 一个Sheet最大允许的行数，设置了之后将输出多个Sheet
    /// </summary>
    public int MaxRowNumberOnASheet { get; set; } = 0;
}
