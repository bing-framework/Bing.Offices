using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Bing.Offices.ApiSnapshot;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 公共 API 基线测试。
/// </summary>
public class PublicApiContractTest
{
    private static readonly IReadOnlyDictionary<string, string> ApiTypeCategories =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.DecoratorAttributeBase"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.FilterAttributeBase"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Attributes.BindFilterAttribute"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelColumnConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingConfigurationMerger"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDiagnostic"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocument"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocumentFactory"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicColumnConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDynamicValidationConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelDynamicColumnMergeMode"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingLayoutConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingStyleConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValidationRuleMergeMode"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingMergeMode"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelValueMappingConfiguration"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportColumnMappingBuilder`2"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExportMappingBuilder`1"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.FluentSetting`2"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IImportMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IExportMappingProfile`1"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfile`2"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileRegistry"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IMappingProfileResolver"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportColumnMappingBuilder`2"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ImportMappingBuilder`1"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingDirection"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingProfileRegistry"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.MappingSourceKind"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelModelAliasRegistry"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.ProfileDescriptor"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelCellValue"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Conversions.ExcelConversionContext"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Configurations.IExcelMappingConfigurationLoader"] = "Execution detail",
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
            ["Bing.Offices.Abstractions:Bing.Offices.ExcelFormat"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnPlacement"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthMode"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelColumnWidthOptions"] = "User API",
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
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.UniqueTracker"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelDynamicMappingColumn"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingLayout"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingColumn"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingPlanFactory"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingSheetPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingStyle"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Providers.IExcelMappingWorkbookPlan"] = "Provider SPI",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelUnsupportedFeaturePolicy"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ValidateMode"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelWhitespacePolicy"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderLineStyle"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderStyle"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyle"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyleReset"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelColor"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelFillPattern"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelHorizontalAlignment"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelVerticalAlignment"] = "Execution detail",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationContext"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationBindingKind"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationBinding"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationRule"] = "User API",
            ["Bing.Offices.Abstractions:Bing.Offices.Validations.INamedExcelValidationRule"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DynamicColumnAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelDateAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelIgnoreAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxLengthAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxValueAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRangeAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRegexAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRequiredAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelUniqueAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.HeaderAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelColumnMappingBuilder`2"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMapping"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingBuilder`1"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.DynamicColumnAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelDateAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelIgnoreAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxLengthAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelMaxValueAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRangeAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRegexAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelRequiredAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ExcelUniqueAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.HeaderAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelColumnMappingBuilder`2"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMapping"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingBuilder`1"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.CsvHelper"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExpressionExtension"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.PropertyInfoExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.TypeExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueMap`1"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider"] = "Provider SPI",
            ["Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.RegexConst"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Styles.Color"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.DuplicationExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.MaxLengthExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.MaxValueExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RangeExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RegexExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RequiredExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.CsvHelper"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExpressionExtension"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.PropertyInfoExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Extensions.TypeExtensions"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueMap`1"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.RegexConst"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Styles.Color"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.DuplicationExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.MaxLengthExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.MaxValueExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RangeExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RegexExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Core:Bing.Offices.Validations.RequiredExcelValidationRule"] = "Execution detail",
            ["Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions"] = "User API",
            ["Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions"] = "Execution detail"
        };

    private static readonly IReadOnlyDictionary<string, ApiMemberGovernancePolicy> ApiMemberGovernancePolicies =
        new Dictionary<string, ApiMemberGovernancePolicy>(StringComparer.Ordinal)
        {
            ["User API"] = new ApiMemberGovernancePolicy(
                "public source and binary contract",
                "preserve; additive changes require compatibility review"),
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

    private sealed class ApiMemberGovernancePolicy
    {
        public ApiMemberGovernancePolicy(string sourceBinaryImpact, string migrationPolicy)
        {
            SourceBinaryImpact = sourceBinaryImpact;
            MigrationPolicy = migrationPolicy;
        }

        public string SourceBinaryImpact { get; }

        public string MigrationPolicy { get; }
    }

    private sealed class ApiMemberGovernanceRecord
    {
        public ApiMemberGovernanceRecord(string memberKey, string category,
            ApiMemberGovernancePolicy policy)
        {
            MemberKey = memberKey;
            Category = category;
            SourceBinaryImpact = policy.SourceBinaryImpact;
            MigrationPolicy = policy.MigrationPolicy;
        }

        public string MemberKey { get; }

        public string Category { get; }

        public string SourceBinaryImpact { get; }

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
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDiagnostic",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocument",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.ExcelMappingDocumentFactory",
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
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ValidateMode",
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
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelColumnMappingBuilder`2",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMapping",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingBuilder`1",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader",
            "Bing.Offices.Core:Bing.Offices.CsvHelper",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExpressionExtension",
            "Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.PropertyInfoExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.TypeExtensions",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueMap`1",
            "Bing.Offices.Core:Bing.Offices.Metadata.MergedRegionInfo",
            "Bing.Offices.Core:Bing.Offices.Metadata.PictureInfo",
            "Bing.Offices.Core:Bing.Offices.Metadata.PictureStyle",
            "Bing.Offices.Core:Bing.Offices.RegexConst",
            "Bing.Offices.Core:Bing.Offices.Styles.Color",
            "Bing.Offices.Core:Bing.Offices.Validations.DateTimeExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.DuplicationExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.ExcelValidationRules",
            "Bing.Offices.Core:Bing.Offices.Validations.MaxLengthExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.MaxValueExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.RangeExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.RegexExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Validations.RequiredExcelValidationRule",
            "Bing.Offices.Npoi:Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.MappingProfileServiceCollectionExtensions"
        };
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly
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
    /// 测试 - provider-neutral 公共成员签名不得引用 NPOI 实现类型。
    /// </summary>
    [Fact]
    public void PublicApi_PublicMembers_ShouldNotExposeNpoiTypes()
    {
        // Arrange
        var assemblies = new[] { typeof(IExcelImporter).Assembly, typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly };

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
            typeof(NpoiExcelImporter).Assembly
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
            new[] { "User API", "Provider SPI", "Compatibility", "Execution detail" }));
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
                new[] { "User API", "Provider SPI", "Compatibility", "Execution detail" });
            Assert.False(string.IsNullOrWhiteSpace(record.MemberKey));
            Assert.False(string.IsNullOrWhiteSpace(record.SourceBinaryImpact));
            Assert.False(string.IsNullOrWhiteSpace(record.MigrationPolicy));
        });
    }

    /// <summary>
    /// 测试 - 发布程序集不得包含指向 Core/NPOI 的生产 InternalsVisibleTo，仅允许测试友元。
    /// </summary>
    [Fact]
    public void PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies()
    {
        // Arrange
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly
        };

        // Act
        var friends = assemblies.SelectMany(assembly => assembly.GetCustomAttributes<InternalsVisibleToAttribute>())
            .Select(attribute => attribute.AssemblyName.Split(',')[0])
            .Where(name => name.IndexOf(".Tests", StringComparison.Ordinal) < 0)
            .ToArray();

        // Assert
        Assert.Empty(friends);
    }

    /// <summary>
    /// 测试 - NPOI 适配程序集只公开 DI 注册入口，不公开实现和低层 NPOI 类型。
    /// </summary>
    [Fact]
    public void PublicApi_NpoiAssembly_ShouldExposeOnlyRegistrationEntry()
    {
        // Arrange
        var assembly = typeof(NpoiExcelImporter).Assembly;

        // Act
        var exported = assembly.GetExportedTypes().Select(type => type.FullName)
            .OrderBy(name => name, StringComparer.Ordinal).ToArray();

        // Assert
        Assert.Equal(new[]
        {
            "Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions"
        }, exported);
        var leaked = assembly.GetExportedTypes().SelectMany(type => type.GetMembers(
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance
                | System.Reflection.BindingFlags.Static))
            .SelectMany(GetSignatureTypes)
            .Where(type => type.FullName?.StartsWith("NPOI.", StringComparison.Ordinal) == true)
            .ToArray();
        Assert.Empty(leaked);
    }

    /// <summary>
    /// 测试 - NPOI 程序集的公开成员签名应精确匹配唯一 DI 注册入口。
    /// </summary>
    [Fact]
    public void PublicApi_NpoiAssembly_ShouldMatchExactMemberBaseline()
    {
        // Arrange
        var expected = new[]
        {
            "type|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions|generic=0",
            "method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddBingOfficesNpoi|static|Microsoft.Extensions.DependencyInjection.IServiceCollection|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0",
        };
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
        Assert.Equal(expected.OrderBy(member => member, StringComparer.Ordinal), actual);
    }

    /// <summary>
    /// 测试 - Abstractions、Core 和 NPOI 程序集的全部公开成员应匹配批准快照。
    /// </summary>
    [Fact]
    public void PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot()
    {
        // Arrange
        var expected = new Dictionary<string, string>(StringComparer.Ordinal)
        {
#if NETCOREAPP3_1
            ["Bing.Offices.Abstractions"] = "7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943",
#elif NET6_0
            ["Bing.Offices.Abstractions"] = "7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943",
#elif NET8_0
            ["Bing.Offices.Abstractions"] = "7F9A2AA819E94B3838097DF2FF374A934CF7F35F3D2E91F3D1DB790F22972943",
#endif
#if NETCOREAPP3_1
            ["Bing.Offices.Core"] = "B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6",
#elif NET6_0
            ["Bing.Offices.Core"] = "B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6",
#elif NET8_0
            ["Bing.Offices.Core"] = "B3661970BBE5AECC06DAD57B1E3F960FA77E70C4D2E66B2DA4910F7823AA2BB6",
#endif
#if NETCOREAPP3_1
            ["Bing.Offices.Npoi"] = "DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE"
#elif NET6_0
            ["Bing.Offices.Npoi"] = "DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE"
#elif NET8_0
            ["Bing.Offices.Npoi"] = "DA163263804A964D8AC2A13D78D6B3858256171CE7729841690FDB56F602CEEE"
#endif
        };
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelMappingConfigurationLoader).Assembly,
            typeof(NpoiExcelImporter).Assembly
        };

        // Act
        var actual = assemblies.ToDictionary(assembly => assembly.GetName().Name,
            GetPublicMemberSnapshotHash, StringComparer.Ordinal);

        // Assert
        foreach (var pair in expected)
        {
            if (!string.Equals(pair.Value, actual[pair.Key], StringComparison.Ordinal))
                throw new InvalidOperationException($"{pair.Key}: expected={pair.Value}; actual={actual[pair.Key]}");
        }
    }

    private static string FormatConstructor(Type type, System.Reflection.ConstructorInfo constructor) =>
        $"constructor|{type.FullName}|{FormatParameters(constructor.GetParameters())}";

    private static string FormatProperty(Type type, System.Reflection.PropertyInfo property) =>
        $"property|{type.FullName}.{property.Name}|{property.PropertyType.FullName}";

    private static string FormatField(Type type, System.Reflection.FieldInfo field) =>
        $"field|{type.FullName}.{field.Name}|{field.FieldType.FullName}";

    private static string FormatMethod(Type type, System.Reflection.MethodInfo method) =>
        $"method|{type.FullName}.{method.Name}|{(method.IsStatic ? "static" : "instance")}|{method.ReturnType.FullName}|"
        + $"{FormatParameters(method.GetParameters())}|generic={method.GetGenericArguments().Length}";

    private static string FormatParameters(IReadOnlyList<System.Reflection.ParameterInfo> parameters) =>
        string.Join(",", parameters.Select(parameter => parameter.ParameterType.FullName));

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

    private static string GetPublicMemberSnapshotHash(System.Reflection.Assembly assembly)
    {
        var directory = Path.GetDirectoryName(assembly.Location)
            ?? throw new InvalidOperationException($"程序集路径不可用: {assembly.FullName}");
        return PublicApiSnapshot.Load(assembly.Location,
            Directory.EnumerateFiles(directory, "*.dll")).Hash;
    }

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

    private static bool IsGovernedProperty(System.Reflection.PropertyInfo property) =>
        property.GetAccessors(true).Any(accessor =>
            accessor.IsPublic || accessor.IsFamily || accessor.IsFamilyOrAssembly);

    private static bool IsGovernedField(System.Reflection.FieldInfo field) =>
        field.IsPublic || field.IsFamily || field.IsFamilyOrAssembly;

    private static IEnumerable<System.Reflection.PropertyInfo> GetGovernedProperties(Type type)
    {
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Public
            | System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.Static
            | System.Reflection.BindingFlags.DeclaredOnly;

        return type.GetProperties(flags).Where(IsGovernedProperty);
    }

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
