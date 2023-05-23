namespace Bing.Offices.Settings;

/// <summary>
/// 工作表设置
/// </summary>
public sealed class SheetSetting : SheetSettingBase
{
    /// <summary>
    /// 工作表名称
    /// </summary>
    private string _sheetName = "Sheet0";

    /// <summary>
    /// 工作表名称
    /// </summary>
    public string SheetName
    {
        get => _sheetName;
        set
        {
            if (!string.IsNullOrWhiteSpace(value))
                _sheetName = value;
        }
    }
}
