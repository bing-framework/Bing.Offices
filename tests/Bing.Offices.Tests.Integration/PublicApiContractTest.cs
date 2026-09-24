using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json;
using Bing.Offices.ApiSnapshot;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 公共 API 基线测试。
/// </summary>
public class PublicApiContractTest
{
    /// <summary>
    /// 按程序集和类型全名索引的公共 API 分类字典。
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> ApiTypeCategories =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.DecoratorAttributeBase"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.FilterAttributeBase"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.BindFilterAttribute"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelColumnConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingConfigurationMerger"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocument"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocumentFactory"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.IO.IFileExportCommitter"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicColumnConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicValidationConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelDynamicColumnMergeMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingLayoutConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingStyleConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValidationRuleMergeMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingMergeMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingConfiguration"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportColumnMappingBuilder`2"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportMappingBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.FluentSetting`2"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IImportMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IExportMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`2"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileRegistry"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileResolver"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportColumnMappingBuilder`2"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportMappingBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingDirection"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingProfileRegistry"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingSourceKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelModelAliasRegistry"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ProfileDescriptor"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellValue"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelConversionContext"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesErrorCode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesOperation"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesStage"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.IBingOfficesExceptionObserver"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesConfigurationException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesImportException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesExportException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesResourceLimitException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesFileCommitException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IExcelMappingConfigurationLoader"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.IExcelValueConverter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.INamedExcelValueConverter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvExportOptions`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvFormulaInjectionPolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportError"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportErrorCode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportOptions`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportResult`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.ICsvExporter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Csv.ICsvImporter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.ExcelFormat"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnPlacement"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthOptions"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelRowHeightOptions"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelCommentConflictPolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelTemplateCellOverwritePolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelComment"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelDynamicColumnDefinition"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelExport"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartAnchor"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartDefinition"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartRange"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartSeries"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartType"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelHeaderCell"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelHeaderRow"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelSheetExportBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelSheetExportRequest"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelUnknownDynamicValuePolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookExportBuilder"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookMetadataOptions"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookExportRequest"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.IExcelExporter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImport"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportError"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportErrorCode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureOptions"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureDiagnostic"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureWorkbookMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportCommentConflictPolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportValidationMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImageMultiplicityPolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImageData"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelNameComparison"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelReadColumnRange"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelRelationRequest"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelResourceLimits"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetSelector"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetSelectorKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportResult"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportRequest"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportRequest`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportResult`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.IExcelImporter"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntity"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityLayoutBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityListRegionBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityLayout`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellBinding`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellBuilder`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityListRegion`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityMergeRegion"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellReference"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellRange"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityImportResult`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityTemplateOptions"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.IExcelEntityExporter"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Entities.IExcelEntityImporter"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.ExcelProviderCapabilities"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelProviderCapabilities"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.UniqueTracker"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelDynamicMappingColumn"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingLayout"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingColumn"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlanFactory"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingSheetPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingStyle"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingWorkbookPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelUnsupportedFeaturePolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelValidationFailureMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWhitespacePolicy"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderLineStyle"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderStyle"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyle"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyleReset"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelColor"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelFillPattern"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelHorizontalAlignment"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelVerticalAlignment"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationContext"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationBindingKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationBinding"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationRule"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.INamedExcelValidationRule"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DynamicColumnAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Dates.ExcelDateOffsetPolicy"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelDateAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelIgnoreAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxLengthAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxValueAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRangeAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRegexAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRequiredAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelUniqueAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.HeaderAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExcelEntityExtensions"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Exceptions.BingOfficesExceptionDispatcher"] = "Provider SPI",
            ["Bing.Offices.Core:Bing.Offices.IO.DefaultFileExportCommitter"] = "Provider SPI",
            ["Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider"] = "Provider SPI",
            ["Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo"] = "Provider User API",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo"] = "Provider User API",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle"] = "Provider User API",
            ["Bing.Offices.Core:Bing.Offices.Styles.Color"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule"] = "Provider SPI",
            ["Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules"] = "Provider SPI",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions"] = "User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.CellExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.RowExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.SheetExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.WorkbookExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.CellStyleExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.FontExtensions"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Exports.NpoiExcelExporter"] = "Provider User API",
            ["Bing.Offices.Npoi:Bing.Offices.Imports.NpoiExcelImporter"] = "Provider User API",
            ["Bing.Offices.MiniExcel:Bing.Offices.Exports.MiniExcelExcelExporter"] = "Provider User API",
            ["Bing.Offices.MiniExcel:Bing.Offices.Imports.MiniExcelExcelImporter"] = "Provider User API",
            ["Bing.Offices.MiniExcel:Bing.Offices.MiniExcel.Extensions.ExcelMiniExcelServiceCollectionExtensions"] = "User API",
            ["Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Exports.ClosedXmlExcelExporter"] = "Provider User API",
            ["Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Imports.ClosedXmlExcelImporter"] = "Provider User API",
            ["Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Extensions.ExcelClosedXmlServiceCollectionExtensions"] = "User API",
            ["Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.ClosedXmlProviderOptions"] = "Provider User API",
        };

    /// <summary>
    /// 按 API 分类索引的成员治理策略字典。
    /// </summary>
    private static readonly IReadOnlyDictionary<string, ApiMemberGovernancePolicy> ApiMemberGovernancePolicies =
        new Dictionary<string, ApiMemberGovernancePolicy>(StringComparer.Ordinal)
        {
            ["User API"] = new ApiMemberGovernancePolicy(
                "public source and binary contract",
                "preserve; additive changes require compatibility review"),
            ["Provider User API"] = new ApiMemberGovernancePolicy(
                "provider-specific public source and binary contract",
                "preserve; additive changes require provider compatibility review"),
            ["Provider SPI"] = new ApiMemberGovernancePolicy(
                "provider source and binary contract",
                "preserve; breaking changes require provider migration and version approval"),
            ["Compatibility"] = new ApiMemberGovernancePolicy(
                "legacy source and binary contract",
                "preserve or obsolete with forwarding and migration guidance"),
            ["Execution detail"] = new ApiMemberGovernancePolicy(
                "currently public implementation surface",
                "do not add dependencies; any visibility or signature change requires API approval")
        };

    /// <summary>
    /// 表示 API 合同测试使用的治理策略。
    /// </summary>
    private sealed class ApiMemberGovernancePolicy
    {
        /// <summary>
        /// 初始化一个 <see cref="ApiMemberGovernancePolicy" /> 类型的实例。
        /// </summary>
        /// <param name="sourceBinaryImpact">成员对源码和二进制兼容性的影响说明。</param>
        /// <param name="migrationPolicy">成员变更的迁移策略说明。</param>
        public ApiMemberGovernancePolicy(string sourceBinaryImpact, string migrationPolicy)
        {
            SourceBinaryImpact = sourceBinaryImpact;
            MigrationPolicy = migrationPolicy;
        }

        /// <summary>
        /// 获取源代码二进制影响。
        /// </summary>
        public string SourceBinaryImpact { get; }

        /// <summary>
        /// 获取迁移策略。
        /// </summary>
        public string MigrationPolicy { get; }
    }

    /// <summary>
    /// 表示 API 合同测试使用的治理记录。
    /// </summary>
    private sealed class ApiMemberGovernanceRecord
    {
        /// <summary>
        /// 初始化一个 <see cref="ApiMemberGovernanceRecord" /> 类型的实例。
        /// </summary>
        /// <param name="memberKey">含程序集、类型和签名的成员标识。</param>
        /// <param name="category">成员所属的 API 分类。</param>
        /// <param name="policy">分类对应的兼容性与迁移策略。</param>
        public ApiMemberGovernanceRecord(string memberKey, string category,
            ApiMemberGovernancePolicy policy)
        {
            MemberKey = memberKey;
            Category = category;
            SourceBinaryImpact = policy.SourceBinaryImpact;
            MigrationPolicy = policy.MigrationPolicy;
        }

        /// <summary>
        /// 获取成员键。
        /// </summary>
        public string MemberKey { get; }

        /// <summary>
        /// 获取分类。
        /// </summary>
        public string Category { get; }

        /// <summary>
        /// 获取源代码二进制影响。
        /// </summary>
        public string SourceBinaryImpact { get; }

        /// <summary>
        /// 获取迁移策略。
        /// </summary>
        public string MigrationPolicy { get; }
    }

    /// <summary>
    /// 测试 - 发布程序集的公开顶层类型应与已批准的 Stream-first API 基线一致。
    /// </summary>
    [Fact]
    public void PublicApi_ReleaseAssemblies_ShouldMatchApprovedBaseline()
    {
        // Arrange
        var expected = new[]
        {
            "Bing.Offices.Abstractions:Bing.Offices.Attributes.DecoratorAttributeBase",
            "Bing.Offices.Abstractions:Bing.Offices.Attributes.FilterAttributeBase",
            "Bing.Offices.Abstractions:Bing.Offices.Attributes.BindFilterAttribute",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelColumnConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingConfigurationMerger",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocument",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocumentFactory",
            "Bing.Offices.Abstractions:Bing.Offices.IO.IFileExportCommitter",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicColumnConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicValidationConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelDynamicColumnMergeMode",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingLayoutConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingStyleConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValidationRuleMergeMode",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingMergeMode",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingConfiguration",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportColumnMappingBuilder`2",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportMappingBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.FluentSetting`2",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`1",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IImportMappingProfile`1",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IExportMappingProfile`1",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`2",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileRegistry",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileResolver",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportColumnMappingBuilder`2",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportMappingBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingDirection",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingProfileRegistry",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingSourceKind",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelModelAliasRegistry",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ProfileDescriptor",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellKind",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellValue",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelConversionContext",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesErrorCode",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesOperation",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesStage",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.IBingOfficesExceptionObserver",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesConfigurationException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesImportException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesExportException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesResourceLimitException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesFileCommitException",
            "Bing.Offices.Abstractions:Bing.Offices.Exceptions.BingOfficesUnsupportedFeatureException",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IExcelMappingConfigurationLoader",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.IExcelValueConverter",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.INamedExcelValueConverter",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvExportOptions`1",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvFormulaInjectionPolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportError",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportErrorCode",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportOptions`1",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportResult`1",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.ICsvExporter",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.ICsvImporter",
            "Bing.Offices.Abstractions:Bing.Offices.ExcelFormat",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnPlacement",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthMode",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelRowHeightOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelCommentConflictPolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelTemplateCellOverwritePolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelComment",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelDynamicColumnDefinition",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelExport",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartAnchor",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartDefinition",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartRange",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartSeries",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelChartType",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelHeaderCell",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelHeaderRow",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelSheetExportBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelSheetExportRequest",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelUnknownDynamicValuePolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookExportBuilder",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookMetadataOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookExportRequest",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.IExcelExporter",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImport",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportError",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportErrorCode",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureDiagnostic",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureWorkbookMode",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportCommentConflictPolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportValidationMode",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImageMultiplicityPolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImageData",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelNameComparison",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelReadColumnRange",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelRelationRequest",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelResourceLimits",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetSelector",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetSelectorKind",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportResult",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelSheetImportRequest",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportRequest`1",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWorkbookImportResult`1",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.IExcelImporter",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntity",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellBinding`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellRange",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityCellReference",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityImportResult`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityLayout`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityLayoutBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityListRegion`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityListRegionBuilder`1",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityMergeRegion",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.ExcelEntityTemplateOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.IExcelEntityExporter",
            "Bing.Offices.Abstractions:Bing.Offices.Entities.IExcelEntityImporter",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.ExcelProviderCapabilities",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelProviderCapabilities",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.UniqueTracker",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelDynamicMappingColumn",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingLayout",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingColumn",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlan",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlanFactory",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingSheetPlan",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingWorkbookPlan",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelUnsupportedFeaturePolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelValidationFailureMode",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWhitespacePolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderLineStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyleReset",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelColor",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelFillPattern",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelHorizontalAlignment",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelVerticalAlignment",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationContext",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationBindingKind",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationBinding",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationRule",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.INamedExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider",
            "Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DynamicColumnAttribute",
            "Bing.Offices.Core:Bing.Offices.Dates.ExcelDateOffsetPolicy",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelDateAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelIgnoreAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxLengthAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxValueAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelRangeAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelRegexAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelRequiredAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ExcelUniqueAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.HeaderAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader",
            "Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExcelEntityExtensions",
            "Bing.Offices.Core:Bing.Offices.Exceptions.BingOfficesExceptionDispatcher",
            "Bing.Offices.Core:Bing.Offices.IO.DefaultFileExportCommitter",
            "Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo",
            "Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo",
            "Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle",
            "Bing.Offices.Core:Bing.Offices.Styles.Color",
            "Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.CellExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.CellStyleExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.FontExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.RowExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.SheetExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.WorkbookExtensions",
            "Bing.Offices.Npoi:Bing.Offices.Exports.NpoiExcelExporter",
            "Bing.Offices.Npoi:Bing.Offices.Imports.NpoiExcelImporter",
            "Bing.Offices.MiniExcel:Bing.Offices.Exports.MiniExcelExcelExporter",
            "Bing.Offices.MiniExcel:Bing.Offices.Imports.MiniExcelExcelImporter",
            "Bing.Offices.MiniExcel:Bing.Offices.MiniExcel.Extensions.ExcelMiniExcelServiceCollectionExtensions",
            "Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Exports.ClosedXmlExcelExporter",
            "Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Imports.ClosedXmlExcelImporter",
            "Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.Extensions.ExcelClosedXmlServiceCollectionExtensions",
            "Bing.Offices.ClosedXml:Bing.Offices.ClosedXml.ClosedXmlProviderOptions",
            "Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions"
        };
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly,
            typeof(MiniExcelExcelImporter).Assembly,
            typeof(ClosedXmlExcelImporter).Assembly
        };

        // Act
        var actual = assemblies.SelectMany(assembly => assembly.GetExportedTypes())
            .Where(type => type.IsPublic)
            .Select(type => $"{type.Assembly.GetName().Name}:{type.FullName}")
            .OrderBy(type => type, StringComparer.Ordinal)
            .ToArray();

        // Assert
        var missing = expected.Except(actual, StringComparer.Ordinal).OrderBy(type => type).ToArray();
        var unexpected = actual.Except(expected, StringComparer.Ordinal).OrderBy(type => type).ToArray();
        Assert.True(missing.Length == 0 && unexpected.Length == 0,
            $"Missing: {string.Join("; ", missing)}\nUnexpected: {string.Join("; ", unexpected)}");
    }

    /// <summary>
    /// 测试 - 公共异常基类只用于统一捕获和派生合同，不允许调用方直接实例化。
    /// </summary>
    [Fact]
    public void BingOfficesException_BaseType_ShouldBeAbstract()
    {
        Assert.True(typeof(BingOfficesException).IsAbstract);
    }

    /// <summary>
    /// 测试 - provider-neutral 公共成员签名不得引用 NPOI 实现类型。
    /// </summary>
    [Fact]
    public void PublicApi_PublicMembers_ShouldNotExposeNpoiTypes()
    {
        // Arrange
        var assemblies = new[] { typeof(IExcelImporter).Assembly, typeof(ExcelMappingConfigurationLoader).Assembly };

        // Act
        var leaked = assemblies.SelectMany(assembly => assembly.GetExportedTypes())
            .SelectMany(type => type.GetMembers(System.Reflection.BindingFlags.Public
                | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static))
            .SelectMany(member => GetSignatureTypes(member))
            .Where(type => type.FullName?.StartsWith("NPOI.", StringComparison.Ordinal) == true)
            .Distinct()
            .ToArray();

        // Assert
        Assert.Empty(leaked);
    }

    /// <summary>
    /// 测试 - 每个发布类型都必须有稳定的 API 分类，Provider SPI 不得意外成为普通用户入口。
    /// </summary>
    [Fact]
    public void PublicApi_ExportedTypes_ShouldHaveGovernedClassification()
    {
        // Arrange
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly,
            typeof(MiniExcelExcelImporter).Assembly,
            typeof(ClosedXmlExcelImporter).Assembly
        };

        // Act
        var exported = assemblies.SelectMany(assembly => assembly.GetExportedTypes())
            .Select(type => new
            {
                Type = type,
                Key = $"{type.Assembly.GetName().Name}:{type.FullName}"
            })
            .ToArray();
        var actualKeys = exported.Select(item => item.Key).ToHashSet(StringComparer.Ordinal);
        var missing = ApiTypeCategories.Keys.Except(actualKeys, StringComparer.Ordinal).ToArray();
        var unexpected = actualKeys.Except(ApiTypeCategories.Keys, StringComparer.Ordinal).ToArray();

        // Assert
        Assert.True(missing.Length == 0 && unexpected.Length == 0,
            $"Missing classifications: {string.Join("; ", missing)}\nUnexpected exported types: {string.Join("; ", unexpected)}");
        Assert.All(exported, item => Assert.Contains(ApiTypeCategories[item.Key],
            new[] { "User API", "Provider User API", "Provider SPI", "Compatibility", "Execution detail" }));
        Assert.All(exported.Where(item => ApiTypeCategories[item.Key] == "Provider SPI"), item =>
            Assert.Equal(EditorBrowsableState.Never,
                item.Type.GetCustomAttribute<EditorBrowsableAttribute>()?.State));

        var memberLedger = new Dictionary<string, ApiMemberGovernanceRecord>(StringComparer.Ordinal);
        foreach (var item in exported)
        {
            var category = ApiTypeCategories[item.Key];
            var policy = ApiMemberGovernancePolicies[category];
            foreach (var member in GetPublicMembers(item.Type))
            {
                var memberKey = FormatMemberKey(item.Type, member);
                Assert.True(memberLedger.TryAdd(memberKey,
                    new ApiMemberGovernanceRecord(memberKey, category, policy)),
                    $"Duplicate public member key: {memberKey}");
            }
        }
        Assert.NotEmpty(memberLedger);
        Assert.All(memberLedger.Values, record =>
        {
            Assert.Contains(record.Category,
                new[] { "User API", "Provider User API", "Provider SPI", "Compatibility", "Execution detail" });
            Assert.False(string.IsNullOrWhiteSpace(record.MemberKey));
            Assert.False(string.IsNullOrWhiteSpace(record.SourceBinaryImpact));
            Assert.False(string.IsNullOrWhiteSpace(record.MigrationPolicy));
        });
    }

    /// <summary>
    /// 测试 - 生产程序集只允许精确批准的职责级测试友元，不得按名称片段放行。
    /// </summary>
    [Fact]
    public void PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies()
    {
        // Arrange
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly,
            typeof(MiniExcelExcelImporter).Assembly,
            typeof(ClosedXmlExcelImporter).Assembly
        };

        // Act
        var approvedFriendsByAssembly = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal)
        {
            [typeof(IExcelImporter).Assembly.GetName().Name] = new HashSet<string>(
                new[] { "Bing.Offices.Tests", "Bing.Offices.Npoi.Tests" }, StringComparer.Ordinal),
            [typeof(ExcelMappingConfigurationLoader).Assembly.GetName().Name] = new HashSet<string>(
                new[] { "Bing.Offices.Tests", "Bing.Offices.Npoi.Tests" }, StringComparer.Ordinal),
            [typeof(NpoiExcelImporter).Assembly.GetName().Name] = new HashSet<string>(
                new[] { "Bing.Offices.Npoi.Tests", "Bing.Offices.Npoi.Tests.Integration" }, StringComparer.Ordinal),
            [typeof(MiniExcelExcelImporter).Assembly.GetName().Name] = new HashSet<string>(
                new[] { "Bing.Offices.MiniExcel.Tests", "Bing.Offices.MiniExcel.Tests.Integration" }, StringComparer.Ordinal),
            [typeof(ClosedXmlExcelImporter).Assembly.GetName().Name] = new HashSet<string>(
                new[] { "Bing.Offices.ClosedXml.Tests", "Bing.Offices.ClosedXml.Tests.Integration" }, StringComparer.Ordinal)
        };

        // Assert
        Assert.Equal(assemblies.Length, approvedFriendsByAssembly.Count);
        foreach (var assembly in assemblies)
        {
            var actual = assembly.GetCustomAttributes<InternalsVisibleToAttribute>()
                .Select(attribute => attribute.AssemblyName.Split(',')[0])
                .ToHashSet(StringComparer.Ordinal);
            Assert.True(approvedFriendsByAssembly[assembly.GetName().Name].SetEquals(actual),
                $"{assembly.GetName().Name} 的 IVT 与逐程序集批准列表不一致。实际：{string.Join(",", actual)}");
            Assert.DoesNotContain(actual, friend => friend.EndsWith(".Tests.Fake", StringComparison.Ordinal));
        }
    }

    /// <summary>
    /// 测试 - NPOI 适配程序集公开批准的扩展入口，但不公开内部实现辅助类型。
    /// </summary>
    [Fact]
    public void PublicApi_NpoiAssembly_ShouldExposeApprovedProviderEntries()
    {
        // Arrange
        var assembly = typeof(NpoiExcelImporter).Assembly;

        // Act
        var exported = assembly.GetExportedTypes().Select(type => type.FullName)
            .OrderBy(name => name, StringComparer.Ordinal).ToArray();

        // Assert
        Assert.Equal(new[]
        {
            "Bing.Offices.Exports.NpoiExcelExporter",
            "Bing.Offices.Imports.NpoiExcelImporter",
            "Bing.Offices.Npoi.Extensions.CellExtensions",
            "Bing.Offices.Npoi.Extensions.CellStyleExtensions",
            "Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions",
            "Bing.Offices.Npoi.Extensions.FontExtensions",
            "Bing.Offices.Npoi.Extensions.RowExtensions",
            "Bing.Offices.Npoi.Extensions.SheetExtensions",
            "Bing.Offices.Npoi.Extensions.WorkbookExtensions"
        }, exported);
        Assert.DoesNotContain(assembly.GetExportedTypes(), type =>
            type.FullName == "Bing.Offices.Npoi.Extensions.InternalExtensions"
            || type.FullName == "Bing.Offices.Npoi.Extensions.PictureTypeResolver");
    }

    /// <summary>
    /// 测试 - NPOI 程序集应公开批准的 Provider User API 入口。
    /// </summary>
    [Fact]
    public void PublicApi_NpoiAssembly_ShouldMatchExactMemberBaseline()
    {
        var assembly = typeof(NpoiExcelImporter).Assembly;

        // Act
        var actual = assembly.GetExportedTypes()
            .OrderBy(type => type.FullName, StringComparer.Ordinal)
            .SelectMany(type => new[]
            {
                $"type|{type.FullName}|generic={type.GetGenericArguments().Length}"
            }.Concat(type.GetConstructors(System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static)
                .Select(constructor => FormatConstructor(type, constructor)))
            .Concat(type.GetProperties(System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                    | System.Reflection.BindingFlags.DeclaredOnly)
                .Select(property => FormatProperty(type, property)))
            .Concat(type.GetFields(System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                    | System.Reflection.BindingFlags.DeclaredOnly)
                .Select(field => FormatField(type, field)))
            .Concat(type.GetMethods(System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                    | System.Reflection.BindingFlags.DeclaredOnly)
                .Where(method => !method.IsSpecialName)
                .Select(method => FormatMethod(type, method)))
            .ToArray());

        actual = actual.OrderBy(member => member, StringComparer.Ordinal).ToArray();

        // Assert
        Assert.Contains(actual, member => member.Contains("ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi",
            StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("CellExtensions.GetStringValue", StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("SheetExtensions.InsertRows", StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("WorkbookExtensions.GetExcelFormat", StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("CellStyleExtensions.SetBorder", StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("FontExtensions.DefaultFont", StringComparison.Ordinal));
        Assert.Contains(actual, member => member.Contains("RowExtensions.GetOrCreateCell", StringComparison.Ordinal));
    }

    /// <summary>
    /// 测试 - API canonicalizer 必须保留泛型参数、参数名、默认值、修饰符和关键 attribute。
    /// </summary>
    [Fact]
    public void PublicApiSnapshot_CanonicalLines_ShouldIncludeGovernedMetadata()
    {
        var abstractions = PublicApiSnapshot.Load(typeof(IExcelImporter).Assembly.Location,
            Directory.EnumerateFiles(Path.GetDirectoryName(typeof(IExcelImporter).Assembly.Location)!, "*.dll"));

        Assert.Contains(
            "type|Bing.Offices.IO.IFileExportCommitter|kind=interface|visibility=public|modifiers=abstract|base=<null>|interfaces=|generic=|attributes=System.ComponentModel.EditorBrowsableAttribute(enum(System.ComponentModel.EditorBrowsableState)=1;)",
            abstractions.Lines);
        Assert.Contains(
            "method|Bing.Offices.Exports.IExcelExporter.ExportToFile|visibility=public|modifiers=instance,abstract,virtual,hidebysig|return=System.Void|returnAttributes=|params=request:value:Bing.Offices.Exports.ExcelWorkbookExportRequest|attributes=,path:value:System.String|attributes=,cancellationToken:value:System.Threading.CancellationToken?=null|attributes=System.Runtime.InteropServices.OptionalAttribute(;)|generic=|attributes=",
            abstractions.Lines);
        Assert.Contains(
            "method|Bing.Offices.Csv.ICsvExporter.ExportToFile|visibility=public|modifiers=instance,abstract,virtual,hidebysig|return=System.Void|returnAttributes=|params=data:value:System.Collections.Generic.IEnumerable`1[[T]]|attributes=,path:value:System.String|attributes=,options:value:Bing.Offices.Csv.CsvExportOptions`1[[T]]?=null|attributes=System.Runtime.InteropServices.OptionalAttribute(;),cancellationToken:value:System.Threading.CancellationToken?=null|attributes=System.Runtime.InteropServices.OptionalAttribute(;)|generic=T:class,new():constraints=|attributes=|attributes=",
            abstractions.Lines);
        Assert.Contains(
            "property|Bing.Offices.Imports.ExcelImportFailureOptions.MaxCopiedPictureBytes|type=System.Nullable`1[[System.Int64]]|params=|accessors=MaxCopiedPictureBytes:public:instance,hidebysig,MaxCopiedPictureBytes:public:instance,hidebysig|attributes=",
            abstractions.Lines);
        Assert.Contains(
            "property|Bing.Offices.Imports.ExcelImportFailureOptions.MaxEstimatedTargetObjects|type=System.Nullable`1[[System.Int64]]|params=|accessors=MaxEstimatedTargetObjects:public:instance,hidebysig,MaxEstimatedTargetObjects:public:instance,hidebysig|attributes=",
            abstractions.Lines);
    }

    /// <summary>
    /// 测试 - Abstractions、Core、NPOI 和 MiniExcel 程序集的全部公开成员应匹配批准快照。
    /// </summary>
    [Fact]
    public void PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot()
    {
        var baselinePath = Path.Combine(FindRepositoryRoot(), "build", "api-snapshot-baseline.json");
        Assert.True(File.Exists(baselinePath),
            $"BLOCKED: API baseline is missing: {baselinePath}");
        using var document = JsonDocument.Parse(File.ReadAllText(baselinePath, System.Text.Encoding.UTF8));
        var root = document.RootElement;
        Assert.Equal("bing.offices.public-api.v2", root.GetProperty("schema").GetString());
        Assert.Equal("2.7.0", root.GetProperty("generatorVersion").GetString());
        Assert.False(string.IsNullOrWhiteSpace(root.GetProperty("approvedBy").GetString()),
            "BLOCKED: API baseline approvedBy is empty.");
        Assert.False(string.IsNullOrWhiteSpace(root.GetProperty("approvedAt").GetString()),
            "BLOCKED: API baseline approvedAt is empty.");
        Assert.Equal("logical-v3", root.GetProperty("candidateIdentity")
            .GetProperty("artifactIdentityFormat").GetString());

        var repositoryRoot = FindRepositoryRoot();
        var expectedTfm = GetCurrentTargetFramework();
        var releaseRoot = Path.Combine(repositoryRoot, "output", "release");
        var releaseDirectories = new[]
        {
            Path.Combine(releaseRoot, "netstandard2.0"),
            Path.Combine(releaseRoot, expectedTfm)
        };
        var assemblyPaths = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [typeof(IExcelImporter).Assembly.GetName().Name!] =
                Path.Combine(releaseRoot, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
            [typeof(ExcelMappingConfigurationLoader).Assembly.GetName().Name!] =
                Path.Combine(releaseRoot, "netstandard2.0", "Bing.Offices.Core.dll"),
            [typeof(NpoiExcelImporter).Assembly.GetName().Name!] =
                Path.Combine(releaseRoot, expectedTfm, "Bing.Offices.Npoi.dll"),
            [typeof(MiniExcelExcelImporter).Assembly.GetName().Name!] =
                Path.Combine(releaseRoot, expectedTfm, "Bing.Offices.MiniExcel.dll"),
            [typeof(ClosedXmlExcelImporter).Assembly.GetName().Name!] =
                Path.Combine(releaseRoot, expectedTfm, "Bing.Offices.ClosedXml.dll")
        };

        Assert.True(root.GetProperty("assemblies").TryGetProperty(expectedTfm, out var expected),
            $"BLOCKED: API baseline is missing {expectedTfm}.");

        var missingPaths = assemblyPaths.Where(pair => !File.Exists(pair.Value))
            .Select(pair => $"{pair.Key}={pair.Value}").ToArray();
        Assert.True(missingPaths.Length == 0,
            $"BLOCKED: unified Release API snapshot input is missing for {expectedTfm}: "
            + string.Join(", ", missingPaths));

        var additionalAssemblyPaths = releaseDirectories
            .Where(Directory.Exists)
            .SelectMany(directory => Directory.EnumerateFiles(directory, "*.dll"))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var actual = assemblyPaths.ToDictionary(
            pair => pair.Key,
            pair => PublicApiSnapshot.Load(pair.Value, additionalAssemblyPaths),
            StringComparer.Ordinal);
        var mismatches = new List<string>();
        foreach (var pair in actual)
        {
            if (!expected.TryGetProperty(pair.Key, out var expectedAssembly))
            {
                mismatches.Add(
                    $"API snapshot mismatch: tfm={expectedTfm}; assembly={pair.Key}; "
                    + $"path={assemblyPaths[pair.Key]}; baseline assembly is missing.");
                continue;
            }

            var expectedHash = expectedAssembly.GetProperty("hash").GetString() ?? string.Empty;
            var expectedMemberCount = expectedAssembly.GetProperty("memberCount").GetInt32();
            var expectedLines = expectedAssembly.GetProperty("lines").EnumerateArray()
                .Select(line => line.GetString() ?? string.Empty).ToArray();
            if (!string.Equals(expectedHash, pair.Value.Hash, StringComparison.Ordinal)
                || expectedMemberCount != pair.Value.MemberCount
                || !expectedLines.SequenceEqual(pair.Value.Lines, StringComparer.Ordinal))
            {
                mismatches.Add(BuildSnapshotMismatchMessage(
                    expectedTfm, pair.Key, assemblyPaths[pair.Key], expectedHash,
                    pair.Value.Hash, expectedMemberCount, pair.Value.MemberCount,
                    expectedLines, pair.Value.Lines));
            }
        }

        Assert.True(mismatches.Count == 0, string.Join(Environment.NewLine, mismatches));
    }

    /// <summary>
    /// 获取当前目标框架。
    /// </summary>
    /// <returns>当前编译目标的 net6.0 或 net8.0 标识。</returns>
    private static string GetCurrentTargetFramework()
    {
#if NET6_0
        return "net6.0";
#else
        return "net8.0";
#endif
    }

    /// <summary>
    /// 查找仓库根目录。
    /// </summary>
    /// <returns>包含解决方案的祖先目录；未找到时返回当前工作目录。</returns>
    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Bing.Offices.sln")))
                return directory.FullName;
            directory = directory.Parent;
        }
        return Directory.GetCurrentDirectory();
    }

    /// <summary>
    /// 格式化构造函数签名。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <param name="constructor">构造函数信息。</param>
    /// <returns>包含声明类型和参数类型的构造函数签名。</returns>
    private static string FormatConstructor(Type type, System.Reflection.ConstructorInfo constructor) =>
        $"constructor|{type.FullName}|{FormatParameters(constructor.GetParameters())}";

    /// <summary>
    /// 格式化属性签名。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <param name="property">属性信息。</param>
    /// <returns>包含声明类型、属性名和属性类型的签名。</returns>
    private static string FormatProperty(Type type, System.Reflection.PropertyInfo property) =>
        $"property|{type.FullName}.{property.Name}|{property.PropertyType.FullName}";

    /// <summary>
    /// 格式化字段签名。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <param name="field">字段信息。</param>
    /// <returns>包含声明类型、字段名和字段类型的签名。</returns>
    private static string FormatField(Type type, System.Reflection.FieldInfo field) =>
        $"field|{type.FullName}.{field.Name}|{field.FieldType.FullName}";

    /// <summary>
    /// 格式化方法签名。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <param name="method">要检查的反射方法。</param>
    /// <returns>包含声明类型、方法名、静态性、返回类型、参数和泛型数量的签名。</returns>
    private static string FormatMethod(Type type, System.Reflection.MethodInfo method) =>
        $"method|{type.FullName}.{method.Name}|{(method.IsStatic ? "static" : "instance")}|{method.ReturnType.FullName}|"
        + $"{FormatParameters(method.GetParameters())}|generic={method.GetGenericArguments().Length}";

    /// <summary>
    /// 格式化方法参数。
    /// </summary>
    /// <param name="parameters">参数集合。</param>
    /// <returns>按参数顺序以逗号连接的类型全名。</returns>
    private static string FormatParameters(IReadOnlyList<System.Reflection.ParameterInfo> parameters) =>
        string.Join(",", parameters.Select(parameter => parameter.ParameterType.FullName));

    /// <summary>
    /// 格式化成员键。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <param name="member">成员信息。</param>
    /// <returns>包含成员种类、声明类型和签名信息的成员键。</returns>
    private static string FormatMemberKey(Type type, System.Reflection.MemberInfo member)
    {
        if (member is System.Reflection.ConstructorInfo constructor)
            return $"constructor|{type.FullName}|{FormatParameters(constructor.GetParameters())}";
        if (member is System.Reflection.PropertyInfo property)
            return $"property|{type.FullName}.{property.Name}|{property.PropertyType.FullName}";
        if (member is System.Reflection.FieldInfo field)
            return $"field|{type.FullName}.{field.Name}|{field.FieldType.FullName}";
        if (member is System.Reflection.MethodInfo method)
            return FormatMethod(type, method);
        throw new InvalidOperationException($"Unsupported public member: {member.MemberType}");
    }

    /// <summary>
    /// 构建快照差异消息。
    /// </summary>
    /// <param name="targetFramework">目标框架。</param>
    /// <param name="assemblyName">程序集名称。</param>
    /// <param name="assemblyPath">目标文件或目录路径。</param>
    /// <param name="expectedHash">预期哈希值。</param>
    /// <param name="actualHash">实际哈希值。</param>
    /// <param name="expectedMemberCount">预期成员数量。</param>
    /// <param name="actualMemberCount">实际成员数量。</param>
    /// <param name="expectedLines">预期行集合。</param>
    /// <param name="actualLines">实际行集合。</param>
    /// <returns>包含目标框架、程序集、哈希、成员数量及增删成员的差异说明。</returns>
    private static string BuildSnapshotMismatchMessage(
        string targetFramework,
        string assemblyName,
        string assemblyPath,
        string expectedHash,
        string actualHash,
        int expectedMemberCount,
        int actualMemberCount,
        IEnumerable<string> expectedLines,
        IEnumerable<string> actualLines)
    {
        var expectedSet = new HashSet<string>(expectedLines, StringComparer.Ordinal);
        var actualSet = new HashSet<string>(actualLines, StringComparer.Ordinal);
        var added = actualSet.Except(expectedSet, StringComparer.Ordinal)
            .OrderBy(line => line, StringComparer.Ordinal);
        var removed = expectedSet.Except(actualSet, StringComparer.Ordinal)
            .OrderBy(line => line, StringComparer.Ordinal);
        return $"API snapshot mismatch: tfm={targetFramework}; assembly={assemblyName}; "
            + $"path={assemblyPath}; expectedHash={expectedHash}; actualHash={actualHash}; "
            + $"expectedMemberCount={expectedMemberCount}; actualMemberCount={actualMemberCount}; "
            + $"added=[{string.Join(" || ", added)}]; removed=[{string.Join(" || ", removed)}]";
    }

    /// <summary>
    /// 获取签名中的类型集合。
    /// </summary>
    /// <param name="member">成员信息。</param>
    /// <returns>成员参数、返回值、属性或字段所引用的类型序列。</returns>
    private static IEnumerable<Type> GetSignatureTypes(System.Reflection.MemberInfo member)
    {
        if (member is System.Reflection.MethodBase method)
        {
            foreach (var parameter in method.GetParameters())
                yield return parameter.ParameterType;
            if (method is System.Reflection.MethodInfo info)
                yield return info.ReturnType;
        }
        if (member is System.Reflection.PropertyInfo property)
            yield return property.PropertyType;
        if (member is System.Reflection.FieldInfo field)
            yield return field.FieldType;
    }

    /// <summary>
    /// 获取类型的公共成员。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>类型自身声明并纳入治理的构造函数、属性、字段和普通方法。</returns>
    private static IEnumerable<System.Reflection.MemberInfo> GetPublicMembers(Type type)
    {
        foreach (var constructor in GetGovernedConstructors(type))
            yield return constructor;
        foreach (var property in GetGovernedProperties(type))
            yield return property;
        foreach (var field in GetGovernedFields(type))
            yield return field;
        foreach (var method in GetGovernedMethods(type))
        {
            if (!method.IsSpecialName)
                yield return method;
        }
    }

    /// <summary>
    /// 获取受管控的构造函数。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>类型自身声明的 public、protected 或 protected internal 构造函数。</returns>
    private static IEnumerable<System.Reflection.ConstructorInfo> GetGovernedConstructors(Type type)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Public
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.Static
            | System.Reflection.BindingFlags.DeclaredOnly;

        return type.GetConstructors(flags).Where(constructor =>
            constructor.IsPublic || constructor.IsFamily || constructor.IsFamilyOrAssembly);
    }

    /// <summary>
    /// 获取受管控的方法。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>类型自身声明的 public、protected 或 protected internal 方法。</returns>
    private static IEnumerable<System.Reflection.MethodInfo> GetGovernedMethods(Type type)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Public
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.Static
            | System.Reflection.BindingFlags.DeclaredOnly;

        return type.GetMethods(flags).Where(method =>
            method.IsPublic || method.IsFamily || method.IsFamilyOrAssembly);
    }

    /// <summary>
    /// 判断属性是否受管控。
    /// </summary>
    /// <param name="property">属性信息。</param>
    /// <returns>任一访问器为 public、protected 或 protected internal 时为 true，否则为 false。</returns>
    private static bool IsGovernedProperty(System.Reflection.PropertyInfo property) =>
        property.GetAccessors(true).Any(accessor =>
            accessor.IsPublic || accessor.IsFamily || accessor.IsFamilyOrAssembly);

    /// <summary>
    /// 判断字段是否受管控。
    /// </summary>
    /// <param name="field">字段信息。</param>
    /// <returns>字段为 public、protected 或 protected internal 时为 true，否则为 false。</returns>
    private static bool IsGovernedField(System.Reflection.FieldInfo field) =>
        field.IsPublic || field.IsFamily || field.IsFamilyOrAssembly;

    /// <summary>
    /// 获取受管控的属性。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>至少一个访问器符合治理可见性要求的自身声明属性。</returns>
    private static IEnumerable<System.Reflection.PropertyInfo> GetGovernedProperties(Type type)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Public
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.Static
            | System.Reflection.BindingFlags.DeclaredOnly;

        return type.GetProperties(flags).Where(IsGovernedProperty);
    }

    /// <summary>
    /// 获取受管控的字段。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>符合治理可见性要求的自身声明字段。</returns>
    private static IEnumerable<System.Reflection.FieldInfo> GetGovernedFields(Type type)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Public
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.Static
            | System.Reflection.BindingFlags.DeclaredOnly;

        return type.GetFields(flags).Where(IsGovernedField);
    }

}
