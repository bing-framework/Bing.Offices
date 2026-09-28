using System.ComponentModel;

namespace Bing.Offices.Entities;

/// <summary>
/// 实体模板输入或输出选项。
/// </summary>
/// <remarks>
/// 导出时模板是输出工作簿的来源；导入时模板仅用于校验工作表和合并结构，实际数据仍从操作的 source 流读取。
/// 当 <see cref="LeaveOpen" /> 为 false 时，Provider 在操作结束后释放模板流。
/// </remarks>
public sealed class ExcelEntityTemplateOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityTemplateOptions" /> 类型的实例。
    /// </summary>
    /// <param name="template">可读取的模板流。</param>
    /// <param name="leaveOpen">完成操作后是否保持模板流打开。</param>
    public ExcelEntityTemplateOptions(Stream template, bool leaveOpen = false)
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        if (!template.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(template));
        Template = template;
        LeaveOpen = leaveOpen;
    }

    /// <summary>
    /// 获取包含既有工作簿模板的可读流。
    /// </summary>
    public Stream Template { get; }
    /// <summary>
    /// 获取操作完成后是否由调用方继续持有模板流。
    /// </summary>
    public bool LeaveOpen { get; }
}
