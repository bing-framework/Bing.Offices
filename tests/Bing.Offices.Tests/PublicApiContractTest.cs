using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 公共 API 基线测试。
/// </summary>
public class PublicApiContractTest
{
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
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.ICellValueConverter",
            "Bing.Offices.Abstractions:Bing.Offices.Configurations.IExcelMappingConfigurationLoader",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.IExcelValueConverter",
            "Bing.Offices.Abstractions:Bing.Offices.Conversions.INamedExcelValueConverter",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvExportOptions`1",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvFormulaInjectionPolicy",
            "Bing.Offices.Abstractions:Bing.Offices.Csv.CsvImportError",
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
            "Bing.Offices.Abstractions:Bing.Offices.Exports.ExcelWorkbookExportRequest",
            "Bing.Offices.Abstractions:Bing.Offices.Exports.IExcelExporter",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImport",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportError",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportErrorCode",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureOptions",
            "Bing.Offices.Abstractions:Bing.Offices.Imports.ExcelImportFailureWorkbookMode",
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
            "Bing.Offices.Abstractions:Bing.Offices.Settings.ExcelSetting",
            "Bing.Offices.Abstractions:Bing.Offices.Settings.SheetSetting",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderLineStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelBorderStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelCellStyle",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelColor",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelFillPattern",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelHorizontalAlignment",
            "Bing.Offices.Abstractions:Bing.Offices.Styles.ExcelVerticalAlignment",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationContext",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.ExcelValidationBindingKind",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationBinding",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.IExcelValidationRule",
            "Bing.Offices.Abstractions:Bing.Offices.Validations.INamedExcelValidationRule",
            "Bing.Offices.Core:Bing.Offices.Attributes.ColumnNameAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DataFormatAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DateTimeAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DecimalScaleAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.DuplicationAttribute",
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
            "Bing.Offices.Core:Bing.Offices.Attributes.MaxLengthAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.MergeColumnsAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.RangeAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.RegexAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.RequiredAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.ValueMappingAttribute",
            "Bing.Offices.Core:Bing.Offices.Attributes.WrapTextAttribute",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelColumnMappingBuilder`2",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMapping",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingBuilder`1",
            "Bing.Offices.Core:Bing.Offices.Configurations.ExcelMappingConfigurationLoader",
            "Bing.Offices.Core:Bing.Offices.Configurations.DefaultExcelMappingConfigurationLoader",
            "Bing.Offices.Core:Bing.Offices.Csv.CsvEntityExporter",
            "Bing.Offices.Core:Bing.Offices.Csv.CsvEntityImporter",
            "Bing.Offices.Core:Bing.Offices.CsvHelper",
            "Bing.Offices.Core:Bing.Offices.Exceptions.OfficeDataConvertException",
            "Bing.Offices.Core:Bing.Offices.Exceptions.OfficeEmptyLineException",
            "Bing.Offices.Core:Bing.Offices.Exceptions.OfficeException",
            "Bing.Offices.Core:Bing.Offices.Exceptions.OfficeHeaderException",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExpressionExtension",
            "Bing.Offices.Core:Bing.Offices.Extensions.CsvStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.ExcelStreamExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.PropertyInfoExtensions",
            "Bing.Offices.Core:Bing.Offices.Extensions.TypeExtensions",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelPropertyMap",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactory",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelTypeMap`1",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelTypeMapFactory",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelValidationBindingFactory",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueMap`1",
            "Bing.Offices.Core:Bing.Offices.Mappings.ExcelValueConverterBindingResolver",
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
            typeof(ExcelTypeMapFactory).Assembly,
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
        var assemblies = new[] { typeof(IExcelImporter).Assembly, typeof(ExcelTypeMapFactory).Assembly,
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
    /// 测试 - 发布程序集不得包含指向 Core/NPOI 的生产 InternalsVisibleTo，仅允许测试友元。
    /// </summary>
    [Fact]
    public void PublicApi_ProductionAssemblies_ShouldNotExposeProductionFriendAssemblies()
    {
        // Arrange
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelTypeMapFactory).Assembly,
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
            "method|Bing.Offices.Npoi.Extensions.ExcelNpoiServiceCollectionExtensions.AddNpoi|static|System.Void|Microsoft.Extensions.DependencyInjection.IServiceCollection|generic=0",
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
#if !NET8_0_OR_GREATER
        return;
#else
        // Arrange
        var expected = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Bing.Offices.Abstractions"] = "8EFC3DFC1FD0F4324C05E67206A55F60FF22A8C1EE86BF3CBFA30A0F78E39F81",
            ["Bing.Offices.Core"] = "40A788EE5B49AF9599942AB68DA946D924FF6062257F1831B5AEBCAA26D760BE",
            ["Bing.Offices.Npoi"] = "A0DBE9808D82547601429D8958C7ED283467031A3763EB9037B19D03F19D80BD"
        };
        var assemblies = new[]
        {
            typeof(IExcelImporter).Assembly,
            typeof(ExcelTypeMapFactory).Assembly,
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
#endif
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

    private static string GetPublicMemberSnapshotHash(System.Reflection.Assembly assembly)
    {
        var lines = new List<string>();
        foreach (var type in assembly.GetExportedTypes().OrderBy(type => type.FullName, StringComparer.Ordinal))
        {
            lines.Add($"type|{type.FullName}|generic={type.GetGenericArguments().Length}");
            foreach (var constructor in type.GetConstructors(System.Reflection.BindingFlags.Public
                         | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static))
                lines.Add(FormatConstructor(type, constructor));
            foreach (var property in type.GetProperties(System.Reflection.BindingFlags.Public
                         | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                         | System.Reflection.BindingFlags.DeclaredOnly))
                lines.Add(FormatProperty(type, property));
            foreach (var field in type.GetFields(System.Reflection.BindingFlags.Public
                         | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                         | System.Reflection.BindingFlags.DeclaredOnly))
                lines.Add(FormatField(type, field));
            foreach (var method in type.GetMethods(System.Reflection.BindingFlags.Public
                         | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Static
                         | System.Reflection.BindingFlags.DeclaredOnly).Where(method => !method.IsSpecialName))
                lines.Add(FormatMethod(type, method));
        }
        var text = string.Join("\n", lines.OrderBy(line => line, StringComparer.Ordinal));
        using var sha256 = System.Security.Cryptography.SHA256.Create();
        var hash = sha256.ComputeHash(System.Text.Encoding.UTF8.GetBytes(text));
        return BitConverter.ToString(hash).Replace("-", string.Empty, StringComparison.Ordinal);
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
}
