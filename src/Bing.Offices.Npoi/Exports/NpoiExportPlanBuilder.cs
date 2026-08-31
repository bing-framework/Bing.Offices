using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Npoi.Internals;

namespace Bing.Offices.Npoi.Exports;

/// <summary>
/// 创建导出工作簿的方向化映射计划，并隔离泛型反射调度。
/// </summary>
internal sealed class NpoiExportPlanBuilder
{
    /// <summary>将请求映射文档编译为不可变工作簿映射计划的工厂。</summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;

    /// <summary>
    /// 初始化导出计划构建器。
    /// </summary>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public NpoiExportPlanBuilder(IExcelMappingPlanFactory mappingPlanFactory)
    {
        _mappingPlanFactory = mappingPlanFactory ?? throw new ArgumentNullException(nameof(mappingPlanFactory));
    }

    /// <summary>
    /// 按规范化请求分组创建导出计划，并将计划映射回原始 Sheet 请求。
    /// </summary>
    /// <param name="request">包含工作表和映射配置的工作簿导出请求。</param>
    /// <returns>每个原始工作表请求对应的不可变列映射计划。</returns>
    public Dictionary<ExcelSheetExportRequest, IExcelMappingPlan> Create(
        ExcelWorkbookExportRequest request)
    {
        var result = new Dictionary<ExcelSheetExportRequest, IExcelMappingPlan>();
        foreach (var group in request.Sheets.GroupBy(GetWorkbookPlanKey, StringComparer.Ordinal))
        {
            var first = group.First();
            var plan = CreateWorkbookPlan(first, group.Select(sheet => sheet.Name).ToArray());
            foreach (var sheet in group)
                result.Add(sheet, plan.Sheets.Single(item => string.Equals(item.Name, sheet.Name,
                    StringComparison.OrdinalIgnoreCase)).Mapping);
        }
        return result;
    }

    /// <summary>生成区分实体类型、映射来源和导出方向的工作簿计划分组键。</summary>
    /// <param name="request">待分组的工作表导出请求。</param>
    /// <returns>可复用映射计划的稳定分组键。</returns>
    private static string GetWorkbookPlanKey(ExcelSheetExportRequest request)
        => NpoiWorkbookPlanKeyBuilder.Create(request.ItemType, request.MappingDocument,
            request.MappingConfiguration, MappingDirection.Export);

    /// <summary>通过反射分派到工作表实体类型对应的泛型计划构建方法。</summary>
    /// <param name="request">包含运行时实体类型的工作表导出请求。</param>
    /// <param name="sheetNames">使用同一映射计划的工作表名称。</param>
    /// <returns>包含各工作表视图的不可变映射计划。</returns>
    private IExcelMappingWorkbookPlan CreateWorkbookPlan(ExcelSheetExportRequest request,
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

    /// <summary>为具体实体类型创建导出方向的工作簿映射计划。</summary>
    /// <typeparam name="T">工作表实体类型。</typeparam>
    /// <param name="request">工作表导出请求。</param>
    /// <param name="sheetNames">使用同一映射计划的工作表名称。</param>
    /// <returns>导出方向的不可变工作簿映射计划。</returns>
    private IExcelMappingWorkbookPlan CreateTypedWorkbookPlan<T>(ExcelSheetExportRequest request,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        return _mappingPlanFactory.CreateWorkbook<T>(request.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, request.MappingConfiguration, MappingDirection.Export, sheetNames);
    }
}
