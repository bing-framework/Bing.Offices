using System;
using System.Collections.Generic;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Testing.Data;
using Bing.Offices.Testing.Models;

namespace Bing.Offices.Testing.Requests;

/// <summary>
/// 创建跨 Provider 合同请求的无 Provider 工厂。
/// </summary>
public static class ContractRequests
{
    /// <summary>
    /// 创建标量导出请求。
    /// </summary>
    /// <returns>包含标量合同数据的导出请求。</returns>
    public static ExcelWorkbookExportRequest ScalarExport() => ExcelExport.Workbook(workbook =>
        workbook.AddSheet("Data", ContractData.ScalarRows()));

    /// <summary>
    /// 创建标量导入请求。
    /// </summary>
    /// <returns>将标量数据导入根工作簿的请求。</returns>
    public static ExcelWorkbookImportRequest<ScalarContractWorkbook> ScalarImport() =>
        ExcelImport.Workbook<ScalarContractWorkbook>(workbook =>
            workbook.Sheet<ScalarContractRow>("Data", root => root.Rows));

    /// <summary>
    /// 创建动态列导出请求。
    /// </summary>
    /// <returns>包含动态区域列的导出请求。</returns>
    public static ExcelWorkbookExportRequest DynamicExport() =>
        ExcelExport.Workbook(workbook => workbook.AddSheet("Data", ContractData.DynamicRows(),
            sheet => sheet.DynamicColumns(row => row.Values, ContractData.DynamicDefinitions())));

    /// <summary>
    /// 创建动态列导入请求。
    /// </summary>
    /// <returns>读取动态区域列的导入请求。</returns>
    public static ExcelWorkbookImportRequest<DynamicContractWorkbook> DynamicImport() =>
        ExcelImport.Workbook<DynamicContractWorkbook>(workbook => workbook
            .Sheet<DynamicContractRow>("Data", root => root.Rows,
                sheet => sheet.DynamicColumns(row => row.Values, ContractData.DynamicDefinitions())));

    /// <summary>
    /// 创建显示值映射配置。
    /// </summary>
    /// <returns>包含显示值映射和列位置的配置。</returns>
    public static ExcelMappingConfiguration MappingConfiguration() => new()
    {
        Columns = new List<ExcelColumnConfiguration>
        {
            new() { PropertyName = nameof(MappingContractRow.Code), Title = "Code" },
            new()
            {
                PropertyName = nameof(MappingContractRow.Name),
                Title = "Display Name",
                ValueMappings = new List<ExcelValueMappingConfiguration>
                {
                    new() { Text = "Alice (display)", Value = "Alice" }
                }
            },
            new() { PropertyName = nameof(MappingContractRow.Quantity), ColumnIndex = 2 }
        }
    };

    /// <summary>
    /// 创建映射导出请求。
    /// </summary>
    /// <returns>使用显示值映射的导出请求。</returns>
    public static ExcelWorkbookExportRequest MappingExport()
    {
        var mapping = MappingConfiguration();
        return ExcelExport.Workbook(workbook => workbook.AddSheet("Data", ContractData.MappingRows(),
            sheet => sheet.Mapping(mapping)));
    }

    /// <summary>
    /// 创建映射导入请求。
    /// </summary>
    /// <returns>使用显示值映射的导入请求。</returns>
    public static ExcelWorkbookImportRequest<MappingContractWorkbook> MappingImport()
    {
        var mapping = MappingConfiguration();
        return ExcelImport.Workbook<MappingContractWorkbook>(workbook => workbook
            .Sheet<MappingContractRow>("Data", root => root.Rows, sheet => sheet.Mapping(mapping)));
    }

    /// <summary>
    /// 创建带命名值转换器的映射配置。
    /// </summary>
    /// <returns>指定命名文本转换器的映射配置。</returns>
    public static ExcelMappingConfiguration MappingConfigurationWithConverter()
    {
        var mapping = MappingConfiguration();
        foreach (var column in mapping.Columns)
            if (column.PropertyName == nameof(MappingContractRow.Name))
                column.ConverterName = "contract-upper";
        return mapping;
    }

    /// <summary>
    /// 创建带命名值转换器的映射导出请求。
    /// </summary>
    /// <returns>同时使用显示值映射和命名转换器的导出请求。</returns>
    public static ExcelWorkbookExportRequest MappingExportWithConverter()
    {
        var mapping = MappingConfigurationWithConverter();
        return ExcelExport.Workbook(workbook => workbook.AddSheet("Data", ContractData.MappingRows(),
            sheet => sheet.Mapping(mapping)));
    }

    /// <summary>
    /// 创建带命名值转换器的映射导入请求。
    /// </summary>
    /// <returns>同时使用显示值映射和命名转换器的导入请求。</returns>
    public static ExcelWorkbookImportRequest<MappingContractWorkbook> MappingImportWithConverter()
    {
        var mapping = MappingConfigurationWithConverter();
        return ExcelImport.Workbook<MappingContractWorkbook>(workbook => workbook
            .Sheet<MappingContractRow>("Data", root => root.Rows, sheet => sheet.Mapping(mapping)));
    }

    /// <summary>
    /// 创建结构化校验错误导出请求。
    /// </summary>
    /// <returns>包含无效测试数据的导出请求。</returns>
    public static ExcelWorkbookExportRequest ValidationExport() => ExcelExport.Workbook(workbook =>
        workbook.AddSheet("Data", ContractData.InvalidValidationRows()));

    /// <summary>
    /// 创建结构化校验错误导入请求。
    /// </summary>
    /// <returns>继续收集校验错误的导入请求。</returns>
    public static ExcelWorkbookImportRequest<ValidationContractWorkbook> ValidationImport() =>
        ExcelImport.Workbook<ValidationContractWorkbook>(workbook => workbook
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 创建关系导出请求。
    /// </summary>
    /// <returns>包含父项和子项工作表的导出请求。</returns>
    public static ExcelWorkbookExportRequest RelationExport() =>
        ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", ContractData.RelationParents())
            .AddSheet("Children", ContractData.RelationChildren()));

    /// <summary>
    /// 创建关系导入请求。
    /// </summary>
    /// <returns>按忽略大小写订单号关联父子项的导入请求。</returns>
    public static ExcelWorkbookImportRequest<RelationContractWorkbook> RelationImport() =>
        ExcelImport.Workbook<RelationContractWorkbook>(workbook =>
        {
            workbook.Sheet("Parents", root => root.Parents);
            workbook.Sheet("Children", root => root.Children);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.OrderNo, child => child.OrderNo,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });
}
