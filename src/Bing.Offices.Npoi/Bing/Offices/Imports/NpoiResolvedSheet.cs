using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Providers;
using Bing.Offices.Internals;

namespace Bing.Offices.Imports;

/// <summary>
/// 已解析的工作表请求执行描述，保存请求与实际物理工作表之间的绑定结果。
/// </summary>
internal sealed class NpoiResolvedSheet
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiResolvedSheet" /> 类型的实例。
    /// </summary>
    /// <param name="request">原始工作表请求。</param>
    /// <param name="index">工作簿中的零基物理索引；未找到时为 -1。</param>
    /// <param name="name">工作簿中的物理工作表名称；未找到时为 null。</param>
    public NpoiResolvedSheet(ExcelSheetImportRequest request, int index, string name)
    {
        Request = request ?? throw new ArgumentNullException(nameof(request));
        Index = index;
        Name = name;
    }

    /// <summary>
    /// 获取原始工作表请求。
    /// </summary>
    public ExcelSheetImportRequest Request { get; }

    /// <summary>
    /// 获取工作簿中的零基物理索引。
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// 获取工作簿中的物理工作表名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取selector 是否成功解析到物理工作表。
    /// </summary>
    public bool Exists => Index >= 0;
}
