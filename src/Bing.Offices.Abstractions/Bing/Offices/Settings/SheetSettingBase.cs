using Bing.Offices.Exports;

namespace Bing.Offices.Settings;

/// <summary>
/// 工作表设置基类
/// </summary>
public abstract class SheetSettingBase
{
    /// <summary>
    /// 起始行索引
    /// </summary>
    private int _startRowIndex = 1;

    /// <summary>
    /// 表头样式
    /// </summary>
    public IBaseStyle HeaderStyle { get; set; } = new CellStyle();

    /// <summary>
    /// 内容样式
    /// </summary>
    public IBaseStyle ColumnStyle { get; set; } = new CellStyle();

    /// <summary>
    /// 启用自适应列宽
    /// </summary>
    public bool AutoColumnWidthEnabled { get; set; }

    /// <summary>
    /// 列头筛选器
    /// </summary>
    public Type ExporterHeaderFilter { get; set; }

    /// <summary>
    /// 列头（集合）筛选器
    /// </summary>
    public Type ExporterHeadersFilter { get; set; }

    /// <summary>
    /// 是否禁用所有筛选器
    /// </summary>
    public bool IsDisableAllFilter { get; set; }

    /// <summary>
    /// 自动居中（设置后全局居中显示）
    /// </summary>
    public bool AutoCenter { get; set; }

    /// <summary>
    /// 起始行索引
    /// </summary>
    public int StartRowIndex
    {
        get => _startRowIndex;
        set
        {
            if (value >= 0)
                _startRowIndex = value;
        }
    }

    /// <summary>
    /// 标题行索引
    /// </summary>
    public int HeaderRowIndex => StartRowIndex - 1;

    /// <summary>
    /// 是否自动索引
    /// </summary>
    public bool AutoIndex { get; set; }
}
