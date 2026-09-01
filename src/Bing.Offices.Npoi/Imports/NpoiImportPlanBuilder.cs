using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Internals;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// 已解析的工作表请求执行描述，保存请求与实际物理工作表之间的绑定结果。
/// </summary>
internal sealed class NpoiResolvedSheet
{
    /// <summary>
    /// 初始化已解析的工作表请求。
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

    /// <summary>获取原始工作表请求。</summary>
    public ExcelSheetImportRequest Request { get; }

    /// <summary>获取工作簿中的零基物理索引。</summary>
    public int Index { get; }

    /// <summary>获取工作簿中的物理工作表名称。</summary>
    public string Name { get; }

    /// <summary>获取 selector 是否成功解析到物理工作表。</summary>
    public bool Exists => Index >= 0;
}

/// <summary>
/// 创建导入工作簿的方向化映射计划，并隔离泛型反射调度。
/// </summary>
internal sealed class NpoiImportPlanBuilder
{
    /// <summary>将请求映射文档编译为不可变工作簿映射计划的工厂。</summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;

    /// <summary>
    /// 初始化导入计划构建器。
    /// </summary>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public NpoiImportPlanBuilder(IExcelMappingPlanFactory mappingPlanFactory)
    {
        _mappingPlanFactory = mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory));
    }

    /// <summary>
    /// 按规范化请求分组创建导入计划，并将计划映射回原始 Sheet 请求。
    /// </summary>
    /// <param name="existingSheets">已成功解析为物理工作表的请求描述。</param>
    /// <returns>每个原始工作表请求对应的不可变列映射计划。</returns>
    public Dictionary<ExcelSheetImportRequest, IExcelMappingPlan> Create(
        IEnumerable<NpoiResolvedSheet> existingSheets)
    {
        var result = new Dictionary<ExcelSheetImportRequest, IExcelMappingPlan>();
        foreach (var group in existingSheets.GroupBy(item => GetWorkbookPlanKey(item.Request), StringComparer.Ordinal))
        {
            var first = group.First().Request;
            var sheetNames = group.Select(item => item.Name).ToArray();
            var plan = CreateWorkbookPlan(first, sheetNames);
            foreach (var item in group)
                result.Add(item.Request, plan.Sheets.Single(sheet => string.Equals(sheet.Name,
                    item.Name, StringComparison.OrdinalIgnoreCase)).Mapping);
        }
        return result;
    }

    /// <summary>生成区分实体类型、映射来源和导入方向的工作簿计划分组键。</summary>
    /// <param name="request">待分组的工作表导入请求。</param>
    /// <returns>可复用映射计划的稳定分组键。</returns>
    private static string GetWorkbookPlanKey(ExcelSheetImportRequest request)
        => NpoiWorkbookPlanKeyBuilder.Create(request.ItemType, request.MappingDocument,
            request.MappingConfiguration, MappingDirection.Import);

    /// <summary>通过反射分派到工作表实体类型对应的泛型计划构建方法。</summary>
    /// <param name="request">包含运行时实体类型的工作表导入请求。</param>
    /// <param name="sheetNames">使用同一映射计划的工作表名称。</param>
    /// <returns>包含各工作表视图的不可变映射计划。</returns>
    private IExcelMappingWorkbookPlan CreateWorkbookPlan(ExcelSheetImportRequest request,
        IReadOnlyList<string> sheetNames)
    {
        var method = GetType().GetMethod(nameof(CreateTypedWorkbookPlan), BindingFlags.Instance
            | BindingFlags.NonPublic).MakeGenericMethod(request.ItemType);
        try
        {
            return (IExcelMappingWorkbookPlan)method.Invoke(this, new object[] { request, sheetNames });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>为具体实体类型创建导入方向的工作簿映射计划。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="request">工作表导入请求。</param>
    /// <param name="sheetNames">使用同一映射计划的工作表名称。</param>
    /// <returns>导入方向的不可变工作簿映射计划。</returns>
    private IExcelMappingWorkbookPlan CreateTypedWorkbookPlan<T>(ExcelSheetImportRequest request,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        return _mappingPlanFactory.CreateWorkbook<T>(request.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, request.MappingConfiguration, MappingDirection.Import, sheetNames);
    }
}
