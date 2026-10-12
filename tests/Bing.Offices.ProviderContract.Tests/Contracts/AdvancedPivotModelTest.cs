using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证高级透视模型的构建校验与不可变快照。
/// </summary>
public sealed class AdvancedPivotModelTest
{
    /// <summary>
    /// 验证旧版单行 Sum 定义仍能通过公开工作簿入口创建请求。
    /// </summary>
    [Fact]
    public void LegacyDefinition_ShouldRemainCompatible()
    {
        var range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1 };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "Legacy",
            SourceSheet = "Data",
            SourceRange = range,
            TargetSheet = "Summary",
            TargetRow = 0,
            TargetColumn = 0,
            RowField = "Region",
            ValueField = "Amount"
        };

        var request = CreateRequest(definition);
        var snapshot = Assert.Single(request.PivotTables);

        Assert.NotSame(definition, snapshot);
        Assert.NotSame(range, snapshot.SourceRange);
        Assert.Equal("Region", snapshot.RowField);
        Assert.Equal("Amount", snapshot.ValueField);
        Assert.Empty(snapshot.RowFields);
        Assert.Empty(snapshot.ColumnFields);
        Assert.Empty(snapshot.ValueFields);
        Assert.Empty(snapshot.Filters);
        Assert.Null(snapshot.MaxExpandedCells);
    }

    /// <summary>
    /// 验证高级字段、筛选去重和请求创建时的深只读快照。
    /// </summary>
    [Fact]
    public void AdvancedDefinition_ShouldSnapshotFieldsAndFilterValues()
    {
        var rowFields = new List<ExcelPivotAxisFieldDefinition>
        {
            new() { Field = "Region" },
            new() { Field = "Category", SortOrder = ExcelPivotSortOrder.Descending,
                SubtotalMode = ExcelPivotSubtotalMode.None }
        };
        var columnFields = new List<ExcelPivotAxisFieldDefinition>
        {
            new() { Field = "Month", SortOrder = ExcelPivotSortOrder.Ascending }
        };
        var valueFields = new List<ExcelPivotValueDefinition>
        {
            new() { Field = "Amount", Caption = "Revenue", NumberFormat = "0.00" },
            new() { Field = "Amount", Caption = "Revenue Count", Aggregation = ExcelPivotAggregation.Count }
        };
        var filterValues = new List<object>
        {
            "Open", "open", 10, 10L, null, null, true, true,
            new DateTime(2026, 10, 5), new DateTime(2026, 10, 5), 2.5m, 2.5d
        };
        var filters = new List<ExcelPivotFilterDefinition>
        {
            new() { Field = "Status", Values = filterValues }
        };
        var definition = CreateAdvanced(rowFields, columnFields, valueFields, filters, maxExpandedCells: 5000);

        var request = CreateRequest(definition);
        rowFields[0] = new ExcelPivotAxisFieldDefinition { Field = "Changed" };
        columnFields.Clear();
        valueFields[0] = new ExcelPivotValueDefinition { Field = "Changed", Caption = "Changed" };
        filterValues[0] = "Closed";
        filterValues.Add("Pending");
        filters[0] = new ExcelPivotFilterDefinition { Field = "Changed", Values = new object[] { "Changed" } };

        var snapshot = Assert.Single(request.PivotTables);
        Assert.Equal(new[] { "Region", "Category" }, snapshot.RowFields.Select(field => field.Field));
        Assert.Equal(new[] { "Month" }, snapshot.ColumnFields.Select(field => field.Field));
        Assert.Equal(new[] { ExcelPivotSortOrder.Default, ExcelPivotSortOrder.Descending },
            snapshot.RowFields.Select(field => field.SortOrder));
        Assert.Equal(new[] { ExcelPivotSubtotalMode.Automatic, ExcelPivotSubtotalMode.None },
            snapshot.RowFields.Select(field => field.SubtotalMode));
        Assert.Equal(new[] { "Revenue", "Revenue Count" }, snapshot.ValueFields.Select(field => field.Caption));
        Assert.Equal(new[] { "Amount", "Amount" }, snapshot.ValueFields.Select(field => field.Field));
        Assert.Equal(ExcelPivotAggregation.Count, snapshot.ValueFields[1].Aggregation);
        Assert.Equal("0.00", snapshot.ValueFields[0].NumberFormat);
        Assert.Equal(5000, snapshot.MaxExpandedCells);
        Assert.Equal(new object[]
        {
            "Open", 10, null, true, new DateTime(2026, 10, 5), 2.5m
        }, snapshot.Filters.Single().Values);

        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotAxisFieldDefinition>)snapshot.RowFields).Add(new ExcelPivotAxisFieldDefinition()));
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotValueDefinition>)snapshot.ValueFields).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotFilterDefinition>)snapshot.Filters).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<object>)snapshot.Filters.Single().Values).Add("Closed"));
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotTableDefinition>)request.PivotTables).Clear());
    }

    /// <summary>
    /// 验证定义校验失败后同一工作簿构建器仍可继续添加有效定义。
    /// </summary>
    [Fact]
    public void InvalidDefinition_ShouldNotPoisonBuilderOrEnumerateSource()
    {
        var enumerated = false;
        var request = ExcelExport.Workbook(builder =>
        {
            builder.AddSheet("Data", Rows(() => enumerated = true));
            Assert.ThrowsAny<ArgumentException>(() => builder.PivotTable(CreateAdvanced(
                new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
                Array.Empty<ExcelPivotAxisFieldDefinition>(),
                new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Region" } },
                Array.Empty<ExcelPivotFilterDefinition>())));
            builder.PivotTable(CreateValidAdvanced());
        });

        Assert.False(enumerated);
        Assert.Single(request.PivotTables);
    }

    /// <summary>
    /// 验证高级模型所有定义边界均在公开 Workbook 构建阶段拒绝。
    /// </summary>
    /// <param name="scenario">待验证的无效模型场景。</param>
    /// <param name="definition">包含该场景的透视定义。</param>
    [Theory]
    [MemberData(nameof(InvalidDefinitions))]
    public void InvalidAdvancedDefinition_ShouldBeRejectedAtBuild(string scenario,
        ExcelPivotTableDefinition definition)
    {
        var exception = Assert.ThrowsAny<ArgumentException>(() => CreateRequest(definition));
        Assert.False(string.IsNullOrWhiteSpace(exception.Message), scenario);
    }

    /// <summary>
    /// 获取高级透视模型的无效定义场景。
    /// </summary>
    public static IEnumerable<object[]> InvalidDefinitions
    {
        get
        {
            var rows = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } };
            var values = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } };
            var filters = new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { "Open" } } };
            var emptyAxes = Array.Empty<ExcelPivotAxisFieldDefinition>();
            var emptyValues = Array.Empty<ExcelPivotValueDefinition>();
            var emptyFilters = Array.Empty<ExcelPivotFilterDefinition>();

            yield return Invalid("null row fields", CreateAdvanced(null, emptyAxes, values, emptyFilters));
            yield return Invalid("null column fields", CreateAdvanced(rows, null, values, emptyFilters));
            yield return Invalid("null value fields", CreateAdvanced(rows, emptyAxes, null, emptyFilters));
            yield return Invalid("null filters", CreateAdvanced(rows, emptyAxes, values, null));
            yield return Invalid("null row member", CreateAdvanced(
                new ExcelPivotAxisFieldDefinition[] { null }, emptyAxes, values, emptyFilters));
            yield return Invalid("null column member", CreateAdvanced(rows,
                new ExcelPivotAxisFieldDefinition[] { null }, values, emptyFilters));
            yield return Invalid("null value member", CreateAdvanced(rows, emptyAxes,
                new ExcelPivotValueDefinition[] { null }, emptyFilters));
            yield return Invalid("null filter member", CreateAdvanced(rows, emptyAxes, values,
                new ExcelPivotFilterDefinition[] { null }));
            yield return Invalid("blank axis field", CreateAdvanced(
                new[] { new ExcelPivotAxisFieldDefinition { Field = " " } }, emptyAxes, values, emptyFilters));
            yield return Invalid("blank value field", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "", Caption = "Total" } }, emptyFilters));
            yield return Invalid("blank caption", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = " " } }, emptyFilters));
            yield return Invalid("blank filter field", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "", Values = Array.Empty<object>() } }));
            yield return Invalid("null filter values", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = null } }));
            yield return Invalid("unknown sort enum", CreateAdvanced(
                new[] { new ExcelPivotAxisFieldDefinition { Field = "Region", SortOrder = (ExcelPivotSortOrder)99 } },
                emptyAxes, values, emptyFilters));
            yield return Invalid("unknown subtotal enum", CreateAdvanced(
                new[] { new ExcelPivotAxisFieldDefinition { Field = "Region", SubtotalMode = (ExcelPivotSubtotalMode)99 } },
                emptyAxes, values, emptyFilters));
            yield return Invalid("unknown aggregation enum", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total", Aggregation = (ExcelPivotAggregation)99 } },
                emptyFilters));
            yield return Invalid("duplicate row axis", CreateAdvanced(
                new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" },
                    new ExcelPivotAxisFieldDefinition { Field = "region" } }, emptyAxes, values, emptyFilters));
            yield return Invalid("duplicate column axis", CreateAdvanced(rows,
                new[] { new ExcelPivotAxisFieldDefinition { Field = "Month" },
                    new ExcelPivotAxisFieldDefinition { Field = "month" } }, values, emptyFilters));
            yield return Invalid("axis shared across row and column", CreateAdvanced(rows,
                new[] { new ExcelPivotAxisFieldDefinition { Field = "region" } }, values, emptyFilters));
            yield return Invalid("axis shared with filter", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "region", Values = Array.Empty<object>() } }));
            yield return Invalid("axis shared with value", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "region", Caption = "Total" } }, emptyFilters));
            yield return Invalid("filter shared with value", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "amount", Values = Array.Empty<object>() } }));
            yield return Invalid("duplicate filter field", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = Array.Empty<object>() },
                    new ExcelPivotFilterDefinition { Field = "status", Values = Array.Empty<object>() } }));
            yield return Invalid("duplicate captions", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" },
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "total" } }, emptyFilters));
            yield return Invalid("caption conflicts with axis", CreateAdvanced(rows, emptyAxes,
                new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "region" } }, emptyFilters));
            yield return Invalid("caption conflicts with filter", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = Array.Empty<object>() },
                    new ExcelPivotFilterDefinition { Field = "total", Values = Array.Empty<object>() } }));
            yield return Invalid("advanced and legacy fields mixed", CreateAdvanced(rows, emptyAxes, values,
                emptyFilters, "Region", "Amount"));
            yield return Invalid("no axes", CreateAdvanced(emptyAxes, emptyAxes, values, emptyFilters));
            yield return Invalid("no values", CreateAdvanced(rows, emptyAxes, emptyValues, emptyFilters));
            yield return Invalid("max expanded cells zero", CreateAdvanced(rows, emptyAxes, values, emptyFilters,
                maxExpandedCells: 0));
            yield return Invalid("max expanded cells negative", CreateAdvanced(rows, emptyAxes, values, emptyFilters,
                maxExpandedCells: -1));
            yield return Invalid("unsupported filter object", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { new object() } } }));
            yield return Invalid("enum is not a numeric filter literal", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { DayOfWeek.Monday } } }));
            yield return Invalid("NaN filter number", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { double.NaN } } }));
            yield return Invalid("infinite filter number", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition { Field = "Status", Values = new object[] { float.PositiveInfinity } } }));
            yield return Invalid("Excel error cell filter object", CreateAdvanced(rows, emptyAxes, values,
                new[] { new ExcelPivotFilterDefinition
                {
                    Field = "Status",
                    Values = new object[] { new ExcelCellValue("#VALUE!", "#VALUE!", ExcelCellKind.Error, errorCode: 15) }
                } }));
        }
    }

    /// <summary>
    /// 创建透视 Workbook 请求以触发公开构建校验和快照。
    /// </summary>
    /// <param name="definition">待验证的透视定义。</param>
    /// <returns>包含透视定义快照的工作簿请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(ExcelPivotTableDefinition definition) =>
        ExcelExport.Workbook(builder => builder.AddSheet("Data", Rows()).PivotTable(definition));

    /// <summary>
    /// 创建默认有效的高级透视定义。
    /// </summary>
    /// <returns>包含一个行轴和一个求和度量的定义。</returns>
    private static ExcelPivotTableDefinition CreateValidAdvanced() => CreateAdvanced(
        new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
        Array.Empty<ExcelPivotAxisFieldDefinition>(),
        new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } },
        Array.Empty<ExcelPivotFilterDefinition>());

    /// <summary>
    /// 使用指定字段集合创建高级透视定义。
    /// </summary>
    /// <param name="rowFields">行轴字段集合。</param>
    /// <param name="columnFields">列轴字段集合。</param>
    /// <param name="valueFields">度量字段集合。</param>
    /// <param name="filters">报表筛选集合。</param>
    /// <param name="legacyRowField">可选的旧版行字段，用于验证模式互斥。</param>
    /// <param name="legacyValueField">可选的旧版求和字段，用于验证模式互斥。</param>
    /// <param name="maxExpandedCells">可选的保守展开单元格上限。</param>
    /// <returns>具有指定描述的透视定义。</returns>
    private static ExcelPivotTableDefinition CreateAdvanced(
        IReadOnlyList<ExcelPivotAxisFieldDefinition> rowFields,
        IReadOnlyList<ExcelPivotAxisFieldDefinition> columnFields,
        IReadOnlyList<ExcelPivotValueDefinition> valueFields,
        IReadOnlyList<ExcelPivotFilterDefinition> filters,
        string legacyRowField = null,
        string legacyValueField = null,
        long? maxExpandedCells = null) => new()
    {
        Name = "AdvancedPivot",
        SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1 },
        TargetSheet = "Summary",
        TargetRow = 0,
        TargetColumn = 0,
        RowField = legacyRowField,
        ValueField = legacyValueField,
        RowFields = rowFields,
        ColumnFields = columnFields,
        ValueFields = valueFields,
        Filters = filters,
        MaxExpandedCells = maxExpandedCells
    };

    /// <summary>
    /// 将无效模型场景和定义打包为理论测试数据。
    /// </summary>
    /// <param name="scenario">场景说明。</param>
    /// <param name="definition">无效透视定义。</param>
    /// <returns>供理论测试使用的参数。</returns>
    private static object[] Invalid(string scenario, ExcelPivotTableDefinition definition) =>
        new object[] { scenario, definition };

    /// <summary>
    /// 创建用于构建请求的业务数据。
    /// </summary>
    /// <param name="onEnumerate">枚举数据时调用的可选观察器。</param>
    /// <returns>包含一条记录的延迟数据序列。</returns>
    private static IEnumerable<PivotModel> Rows(Action onEnumerate = null)
    {
        onEnumerate?.Invoke();
        yield return new PivotModel { Region = "North", Amount = 10 };
    }

    /// <summary>
    /// 透视模型测试使用的最小业务数据行。
    /// </summary>
    private sealed class PivotModel
    {
        /// <summary>
        /// 初始化一个 <see cref="PivotModel" /> 类型的实例。
        /// </summary>
        public PivotModel()
        {
        }

        /// <summary>
        /// 获取或设置分组区域。
        /// </summary>
        public string Region { get; set; }

        /// <summary>
        /// 获取或设置业务金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
