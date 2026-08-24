using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Review Fix 回归测试。
/// </summary>
public sealed class ReviewFixRegressionTest
{
    /// <summary>
    /// 测试 - Document 与 Request 应按 Request 优先级合并，并在构建后保持快照隔离。
    /// </summary>
    [Fact]
    public void MappingSources_RequestOverridesDocumentAndSnapshotsInputs()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = Configuration("文档标题"),
            Export = Configuration("文档导出标题")
        };
        var requestConfiguration = Configuration("请求标题");
        var request = ExcelImport.Workbook<ReviewWorkbook>(builder =>
            builder.Sheet("Data", workbook => workbook.Rows, sheet =>
                sheet.Mapping(document).Mapping(requestConfiguration)));

        // Act
        document.Import.Columns[0].Title = "外部修改文档";
        requestConfiguration.Columns[0].Title = "外部修改请求";

        // Assert
        Assert.Equal("请求标题", request.Sheets[0].MappingConfiguration.Columns[0].Title);
    }

    /// <summary>
    /// 测试 - JSON/XML writer 输出的 v2 文档应能重新加载为等价模型。
    /// </summary>
    [Fact]
    public void MappingDocument_Writers_ShouldRoundTrip()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Profile = "orders",
            ModelAlias = "order-row",
            Import = Configuration("输入"),
            Export = Configuration("输出")
        };

        // Act
        var json = ExcelMappingConfigurationLoader.ToJson(document);
        var xml = ExcelMappingConfigurationLoader.ToXml(document);
        var jsonRoundTrip = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var xmlRoundTrip = ExcelMappingConfigurationLoader.FromXmlDocument(xml);

        // Assert
        Assert.Equal("order-row", jsonRoundTrip.ModelAlias);
        Assert.Equal("输入", jsonRoundTrip.Import.Columns[0].Title);
        Assert.Equal("输出", xmlRoundTrip.Export.Columns[0].Title);
    }

    /// <summary>
    /// 测试 - v2 文档的租户、动态列、样式和布局应在 JSON/XML 往返中保持一致。
    /// </summary>
    [Fact]
    public void MappingDocument_V2ExtendedSchema_ShouldRoundTrip()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            TenantId = "tenant-a",
            ConfigurationVersion = "2026-08-22",
            Import = new ExcelMappingConfiguration
            {
                DynamicColumns =
                {
                    new ExcelMappingDynamicColumnConfiguration
                    {
                        Key = "region", Title = "区域", DataTypeName = "string", ConverterName = "text",
                        ValidationRuleNames = new List<string> { "required", "regex" },
                        ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                        {
                            new() { Name = "regex", Pattern = "^CN-", IgnoreEmpty = false }
                        }
                    }
                },
                Style = new ExcelMappingStyleConfiguration { HeaderStyleKey = "header" },
                Layout = new ExcelMappingLayoutConfiguration { ColumnIndex = 2, PlacementKey = "after-code" }
            }
        };

        // Act
        var json = ExcelMappingConfigurationLoader.ToJson(document);
        var xml = ExcelMappingConfigurationLoader.ToXml(document);
        var jsonRoundTrip = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var xmlRoundTrip = ExcelMappingConfigurationLoader.FromXmlDocument(xml);

        // Assert
        Assert.Equal("tenant-a", jsonRoundTrip.TenantId);
        Assert.Equal("2026-08-22", xmlRoundTrip.ConfigurationVersion);
        Assert.Equal("region", Assert.Single(jsonRoundTrip.Import.DynamicColumns).Key);
        Assert.Equal(new[] { "required", "regex" },
            Assert.Single(jsonRoundTrip.Import.DynamicColumns).ValidationRuleNames);
        Assert.Equal(new[] { "required", "regex" },
            Assert.Single(xmlRoundTrip.Import.DynamicColumns).ValidationRuleNames);
        Assert.Equal("^CN-", Assert.Single(jsonRoundTrip.Import.DynamicColumns).ValidationRules[0].Pattern);
        Assert.False(Assert.Single(xmlRoundTrip.Import.DynamicColumns).ValidationRules[0].IgnoreEmpty);
        Assert.Equal("header", xmlRoundTrip.Import.Style.HeaderStyleKey);
        Assert.Equal(2, xmlRoundTrip.Import.Layout.ColumnIndex);
    }

    /// <summary>
    /// 测试 - XML 未知嵌套节点应报告真实路径而不是固定根节点。
    /// </summary>
    [Fact]
    public void MappingDocument_UnknownNestedXml_ShouldReportExactPath()
    {
        // Arrange
        const string xml = "<ExcelMappingDocument><Version>2</Version><Import><Columns><ExcelColumnConfiguration><Unknown /></ExcelColumnConfiguration></Columns></Import><Export><Columns /></Export></ExcelMappingDocument>";

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() =>
            ExcelMappingConfigurationLoader.FromXmlDocument(xml));

        // Assert
        Assert.Contains("/ExcelMappingDocument/Import/Columns/ExcelColumnConfiguration/Unknown", exception.Message);
    }

    /// <summary>
    /// 测试 - CLR 类型名和未注册业务别名必须在配置边界被拒绝。
    /// </summary>
    [Fact]
    public void MappingDocument_AliasValidation_ShouldRejectClrAndUnknownAliases()
    {
        // Arrange
        const string clr = "{\"version\":2,\"modelAlias\":\"System.String, System.Private.CoreLib\",\"import\":{\"columns\":[]},\"export\":{\"columns\":[]}}";
        var aliases = new ExcelModelAliasRegistry().Register("known-row");
        const string unknown = "{\"version\":2,\"modelAlias\":\"unknown-row\",\"import\":{\"columns\":[]},\"export\":{\"columns\":[]}}";

        // Act / Assert
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(clr));
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(unknown, aliases));
    }

    /// <summary>
    /// 测试 - 已批准的 modelAlias 应绑定模型类型并进入不可变 Plan 身份。
    /// </summary>
    [Fact]
    public void MappingPlan_DocumentIdentity_ShouldResolveApprovedAlias()
    {
        // Arrange
        var aliases = new ExcelModelAliasRegistry().Register("review-row", typeof(ReviewRow), "review");
        var document = new ExcelMappingDocument
        {
            Profile = "review",
            ModelAlias = "review-row",
            Import = Configuration("输入")
        };
        var factory = new ExcelMappingPlanFactory(modelAliases: aliases);

        // Act
        var plan = factory.Create<ReviewRow>(document, null, MappingDirection.Import);

        // Assert
        Assert.Equal("review", plan.ProfileName);
        Assert.Equal("review-row", plan.ModelAlias);
        Assert.Throws<InvalidOperationException>(() =>
            factory.Create<InvalidRegexRow>(document, null, MappingDirection.Import));
    }

    /// <summary>
    /// 测试 - Workbook/Sheet Plan 应复用同一不可变列计划，并按租户和配置版本隔离缓存。
    /// </summary>
    [Fact]
    public void MappingPlan_WorkbookAndTenantVersion_ShouldBeImmutableAndIsolated()
    {
        // Arrange
        var factory = new ExcelMappingPlanFactory();
        var firstDocument = new ExcelMappingDocument
        {
            TenantId = "tenant-a",
            ConfigurationVersion = "v1",
            Import = Configuration("编码")
        };
        var secondDocument = new ExcelMappingDocument
        {
            TenantId = "tenant-b",
            ConfigurationVersion = "v1",
            Import = Configuration("编码")
        };

        // Act
        var first = factory.Create<ReviewRow>(firstDocument, MappingDirection.Import);
        var second = factory.Create<ReviewRow>(secondDocument, MappingDirection.Import);
        var workbook = factory.CreateWorkbook<ReviewRow>(firstDocument, MappingDirection.Import,
            new[] { "Orders", "Archive" });

        // Assert
        Assert.NotSame(first, second);
        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Same(first, workbook.Sheets[0].Mapping);
        Assert.Equal("Orders", workbook.Sheets[0].Name);
        Assert.Equal("Archive", workbook.Sheets[1].Name);
    }

    /// <summary>
    /// 测试 - Plan 构建后修改动态配置集合不应改变既有动态列快照。
    /// </summary>
    [Fact]
    public void MappingPlan_DynamicConfigurationMutation_ShouldNotChangeExistingPlan()
    {
        // Arrange
        var configuration = new ExcelMappingConfiguration
        {
            DynamicColumns =
            {
                new ExcelMappingDynamicColumnConfiguration
                {
                    Key = "region", Title = "区域", DataTypeName = "string",
                    Aliases = new List<string> { "旧区域" },
                    ValidationRuleNames = new List<string> { "rule-a" }
                }
            }
        };
        var document = new ExcelMappingDocument { Import = configuration };
        var rule = new TestNamedRule("rule-a");
        var factory = new ExcelMappingPlanFactory(namedValidationRules: new[] { rule });

        // Act
        var plan = factory.Create<DynamicReviewRow>(document, MappingDirection.Import);
        configuration.DynamicColumns[0].Title = "外部修改";
        configuration.DynamicColumns[0].Aliases[0] = "外部别名";
        configuration.DynamicColumns[0].ValidationRuleNames.Add("rule-b");

        // Assert
        var dynamic = Assert.Single(plan.DynamicColumns);
        Assert.Equal("区域", dynamic.Title);
        Assert.Equal("旧区域", Assert.Single(dynamic.Aliases));
        Assert.Equal(new[] { "rule-a" }, dynamic.ValidationRuleNames);
        Assert.Single(dynamic.ValidationBindings);
    }

    /// <summary>
    /// 测试 - 同一不可变 Plan 可被并发读取并保持引用一致。
    /// </summary>
    [Fact]
    public void MappingPlan_ConcurrentReads_ShouldReuseSamePlan()
    {
        // Arrange
        var factory = new ExcelMappingPlanFactory();
        var document = new ExcelMappingDocument { Import = Configuration("并发") };

        // Act
        var plans = System.Threading.Tasks.Task.WhenAll(
            Enumerable.Range(0, 32).Select(_ => System.Threading.Tasks.Task.Run(() =>
                factory.Create<ReviewRow>(document, MappingDirection.Import)))).GetAwaiter().GetResult();

        // Assert
        Assert.All(plans, plan => Assert.Same(plans[0], plan));
    }

    /// <summary>
    /// 测试 - 有界 production cache 淘汰后应重建被淘汰的 Plan，而不是返回错误租户快照。
    /// </summary>
    [Fact]
    public void MappingPlan_CacheEviction_ShouldRebuildProductionPlan()
    {
        // Arrange
        var factory = new ExcelMappingPlanFactory(cacheCapacity: 1);
        var firstDocument = new ExcelMappingDocument { TenantId = "tenant-a", Import = Configuration("A") };
        var secondDocument = new ExcelMappingDocument { TenantId = "tenant-b", Import = Configuration("B") };

        // Act
        var first = factory.Create<ReviewRow>(firstDocument, MappingDirection.Import);
        _ = factory.Create<ReviewRow>(secondDocument, MappingDirection.Import);
        var rebuilt = factory.Create<ReviewRow>(firstDocument, MappingDirection.Import);

        // Assert
        Assert.NotSame(first, rebuilt);
        Assert.Equal("A", Assert.Single(rebuilt.Columns).Title);
    }

    /// <summary>
    /// 测试 - v1 JSON/XML 归一化时应返回非阻断迁移诊断。
    /// </summary>
    [Fact]
    public void MappingDocument_V1Migration_ShouldReturnDiagnostic()
    {
        // Arrange
        const string json = "{\"columns\":[]}";
        const string xml = "<ExcelMappingConfiguration><Columns /></ExcelMappingConfiguration>";

        // Act
        var jsonDocument = ExcelMappingConfigurationLoader.FromJsonDocument(json, out var jsonDiagnostics);
        var xmlDocument = ExcelMappingConfigurationLoader.FromXmlDocument(xml, out var xmlDiagnostics);

        // Assert
        Assert.Equal(2, jsonDocument.Version);
        Assert.Equal(2, xmlDocument.Version);
        Assert.Contains(jsonDiagnostics, diagnostic => diagnostic.Code == "V1_MIGRATED"
            && diagnostic.Path == "$");
        Assert.Contains(xmlDiagnostics, diagnostic => diagnostic.Code == "V1_MIGRATED"
            && diagnostic.Path == "/ExcelMappingConfiguration");
    }

    /// <summary>
    /// 测试 - 非法正则表达式应在类型映射构建阶段失败，而不是延迟到数据行执行。
    /// </summary>
    [Fact]
    public void MappingPlan_InvalidRegex_ShouldFailAtBuildTime()
    {
        // Act
        var action = () => ExcelTypeMapFactory.Get<InvalidRegexRow>();

        // Assert
        Assert.ThrowsAny<ArgumentException>(action);
    }

    /// <summary>
    /// 测试 - 唯一值 journal 回滚后不得泄漏失败行的首次行号。
    /// </summary>
    [Fact]
    public void UniqueTracker_Rollback_ShouldDiscardFirstRowMetadata()
    {
        // Arrange
        var committed = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var tracker = new UniqueTracker(committed, comparer: StringComparer.OrdinalIgnoreCase);

        // Act
        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 4));
        tracker.RollbackRow();
        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 5));
        tracker.CommitRow();

        // Assert
        Assert.True(tracker.TryGetFirstRowNumber("Code", "A", out var firstRow));
        Assert.Equal(5, firstRow);
        Assert.False(tracker.TryReserve("Code", "A", false, false, 6));
    }

    /// <summary>
    /// 测试 - IgnoreNull 与 IgnoreEmpty 应分别控制 null、空字符串和空白字符串是否参与唯一性跟踪。
    /// </summary>
    [Fact]
    public void UniqueTracker_IgnoreNullAndEmpty_ShouldFollowFlags()
    {
        // Arrange
        var committed = new Dictionary<string, HashSet<string>>();
        var tracker = new UniqueTracker(committed);

        // Act
        tracker.BeginRow();
        var ignoredNull = tracker.TryReserve("Code", null, true, false, 1);
        var trackedNull = tracker.TryReserve("Code", null, false, false, 1);
        var ignoredEmpty = tracker.TryReserve("Code", string.Empty, false, true, 1);
        var trackedWhitespace = tracker.TryReserve("Code", " ", false, false, 1);
        tracker.CommitRow();

        // Assert
        Assert.True(ignoredNull);
        Assert.True(trackedNull);
        Assert.True(ignoredEmpty);
        Assert.True(trackedWhitespace);
        Assert.Equal(2, tracker.TrackedValueCount);
    }

    /// <summary>
    /// 测试 - 配置的 comparer 应同时作用于 pending 和 committed 唯一值。
    /// </summary>
    [Fact]
    public void UniqueTracker_Comparer_ShouldBeUsedForDuplicateDetection()
    {
        // Arrange
        var committed = new Dictionary<string, HashSet<string>>(
            StringComparer.OrdinalIgnoreCase)
        {
            ["Code"] = new HashSet<string>(new[] { "A" }, StringComparer.OrdinalIgnoreCase)
        };
        var tracker = new UniqueTracker(committed, comparer: StringComparer.OrdinalIgnoreCase);

        // Act
        tracker.BeginRow();
        var committedDuplicate = tracker.TryReserve("Code", "a", false, false, 2);
        var first = tracker.TryReserve("Code", "B", false, false, 2);
        var pendingDuplicate = tracker.TryReserve("Code", "b", false, false, 2);

        // Assert
        Assert.False(committedDuplicate);
        Assert.True(first);
        Assert.False(pendingDuplicate);
    }

    /// <summary>
    /// 测试 - 最大跟踪数量只允许达到上限，超出时应立即失败并保留已有状态。
    /// </summary>
    [Fact]
    public void UniqueTracker_MaxTrackedValues_ShouldRejectOnlyOverflow()
    {
        // Arrange
        var tracker = new UniqueTracker(
            new Dictionary<string, HashSet<string>>(), maxTrackedValues: 1);

        // Act
        tracker.BeginRow();
        Assert.True(tracker.TryReserve("Code", "A", false, false, 1));
        tracker.CommitRow();
        Action action = () => tracker.TryReserve("Code", "B", false, false, 2);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
        Assert.Equal(1, tracker.TrackedValueCount);
        Assert.False(tracker.TryGetFirstRowNumber("Code", "B", out _));
    }

    /// <summary>
    /// 测试 - Provider SPI 不得暴露 Expression、Delegate 或 PropertyInfo。
    /// </summary>
    [Fact]
    public void ProviderPlanContract_ShouldExposeOnlyProviderNeutralMembers()
    {
        // Arrange
        var forbidden = new[]
        {
            typeof(Delegate), typeof(System.Linq.Expressions.Expression), typeof(System.Reflection.MemberInfo),
            typeof(Attribute), typeof(Type)
        };

        // Act
        var contract = typeof(IExcelMappingColumn);
        var members = contract.GetMembers();

        // Assert
        Assert.DoesNotContain(members, member => member switch
        {
            System.Reflection.PropertyInfo property => forbidden.Any(type => type.IsAssignableFrom(property.PropertyType)),
            System.Reflection.MethodInfo method => forbidden.Any(type => type.IsAssignableFrom(method.ReturnType)
                || method.GetParameters().Any(parameter => type.IsAssignableFrom(parameter.ParameterType))),
            System.Reflection.FieldInfo field => forbidden.Any(type => type.IsAssignableFrom(field.FieldType)),
            _ => false
        });
        Assert.DoesNotContain(members, member => member is System.Reflection.MethodInfo method
            && (method.Name == "GetValue" || method.Name == "SetValue"));
        Assert.Empty(contract.GetEvents());
        Assert.Empty(contract.GetFields());
        Assert.DoesNotContain(members, member => member is System.Reflection.MethodInfo method
            && method.GetGenericArguments().Length > 0);
        Assert.DoesNotContain(members, member => member is System.Reflection.PropertyInfo property
            && property.PropertyType.IsGenericType
            && property.PropertyType.GetGenericArguments().Any(argument => forbidden.Any(type =>
                type.IsAssignableFrom(argument))));
    }

    /// <summary>
    /// 测试 - Provider SPI 应隐藏于 IntelliSense，旧 object-profile 入口应有迁移标记。
    /// </summary>
    [Fact]
    public void ProviderSpiAndCompatibilityOverloads_ShouldHaveMigrationMetadata()
    {
        // Arrange
        var spiTypes = new[]
        {
            typeof(IExcelMappingPlan), typeof(IExcelMappingColumn), typeof(IExcelMappingPlanFactory),
            typeof(IExcelMappingWorkbookPlan), typeof(IExcelMappingSheetPlan),
            typeof(IExcelMappingStyle), typeof(IExcelMappingLayout), typeof(IExcelDynamicMappingColumn)
        };

        // Act / Assert
        Assert.All(spiTypes, type =>
        {
            var attribute = type.GetCustomAttributes(typeof(System.ComponentModel.EditorBrowsableAttribute), false)
                .Cast<System.ComponentModel.EditorBrowsableAttribute>().SingleOrDefault();
            Assert.NotNull(attribute);
            Assert.Equal(System.ComponentModel.EditorBrowsableState.Never, attribute.State);
        });
        Assert.NotNull(typeof(ExcelMappingDocumentFactory).GetMethod("Create",
            new[] { typeof(object), typeof(ExcelMappingDocument), typeof(ExcelMappingConfiguration),
                typeof(MappingDirection) }).GetCustomAttributes(typeof(ObsoleteAttribute), false).SingleOrDefault());
        Assert.NotNull(typeof(ExcelMappingPlanFactory).GetMethod("Create",
            new[] { typeof(object), typeof(ExcelMappingConfiguration), typeof(MappingDirection) })
            .GetCustomAttributes(typeof(ObsoleteAttribute), false).SingleOrDefault());
    }

    /// <summary>
    /// 测试 - Core 默认注册应提供计划工厂，且 AddNpoi 不得覆盖调用方预注册的替换实现。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_DiDefaultAndReplacement_ShouldPreserveOwnershipBoundary()
    {
        // Arrange
        var defaultServices = new ServiceCollection();
        defaultServices.AddNpoi();
        using var defaultProvider = defaultServices.BuildServiceProvider();
        var replacement = new ExcelMappingPlanFactory(cacheCapacity: 3);
        var replacementServices = new ServiceCollection();
        replacementServices.AddSingleton<IExcelMappingPlanFactory>(replacement);

        // Act
        replacementServices.AddNpoi();
        using var replacementProvider = replacementServices.BuildServiceProvider();

        // Assert
        Assert.IsType<ExcelMappingPlanFactory>(defaultProvider.GetRequiredService<IExcelMappingPlanFactory>());
        Assert.Same(replacement, replacementProvider.GetRequiredService<IExcelMappingPlanFactory>());
    }

    private static ExcelMappingConfiguration Configuration(string title) => new()
    {
        Columns =
        {
            new ExcelColumnConfiguration { PropertyName = nameof(ReviewRow.Name), Title = title }
        }
    };

    private sealed class ReviewWorkbook
    {
        public List<ReviewRow> Rows { get; } = new();
    }

    private sealed class InvalidRegexRow
    {
        [ExcelRegex("[")]
        public string Code { get; set; }
    }

    private sealed class ReviewRow
    {
        public string Name { get; set; }
    }

    private sealed class DynamicReviewRow
    {
        [Bing.Offices.Attributes.DynamicColumn]
        public IDictionary<string, object> Values { get; set; }
    }

    private sealed class TestNamedRule : INamedExcelValidationRule
    {
        public TestNamedRule(string name) => Name = name;
        public string Name { get; }
        public string ErrorMessage => string.Empty;
        public bool Validate(ExcelValidationContext context) => true;
    }
}
