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
/// 创建导入工作簿的方向化映射计划，并隔离泛型反射调度。
/// </summary>
internal sealed class NpoiImportPlanBuilder
{
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
    public Dictionary<ExcelSheetImportRequest, IExcelMappingPlan> Create<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request, IWorkbook workbook,
        IEnumerable<KeyValuePair<ExcelSheetImportRequest, int>> existingSheets)
        where TWorkbook : class, new()
    {
        var result = new Dictionary<ExcelSheetImportRequest, IExcelMappingPlan>();
        foreach (var group in existingSheets.GroupBy(item => GetWorkbookPlanKey(item.Key), StringComparer.Ordinal))
        {
            var first = group.First().Key;
            var sheetNames = group.Select(item => workbook.GetSheetName(item.Value)).ToArray();
            var plan = CreateWorkbookPlan(first, sheetNames);
            foreach (var item in group)
                result.Add(item.Key, plan.Sheets.Single(sheet => string.Equals(sheet.Name,
                    workbook.GetSheetName(item.Value), StringComparison.OrdinalIgnoreCase)).Mapping);
        }
        return result;
    }

    private static string GetWorkbookPlanKey(ExcelSheetImportRequest request)
        => NpoiWorkbookPlanKeyBuilder.Create(request.ItemType, request.MappingDocument,
            request.MappingConfiguration, MappingDirection.Import);

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

    private IExcelMappingWorkbookPlan CreateTypedWorkbookPlan<T>(ExcelSheetImportRequest request,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        return _mappingPlanFactory.CreateWorkbook<T>(request.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, request.MappingConfiguration, MappingDirection.Import, sheetNames);
    }
}
