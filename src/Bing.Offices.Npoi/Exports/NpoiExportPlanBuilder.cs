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

    private static string GetWorkbookPlanKey(ExcelSheetExportRequest request)
        => NpoiWorkbookPlanKeyBuilder.Create(request.ItemType, request.MappingDocument,
            request.MappingConfiguration, MappingDirection.Export);

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

    private IExcelMappingWorkbookPlan CreateTypedWorkbookPlan<T>(ExcelSheetExportRequest request,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        return _mappingPlanFactory.CreateWorkbook<T>(request.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, request.MappingConfiguration, MappingDirection.Export, sheetNames);
    }
}
