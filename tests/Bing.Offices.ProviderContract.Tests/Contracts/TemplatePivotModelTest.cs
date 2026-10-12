using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证模板透视保留和更新请求的公开构建契约。
/// </summary>
public sealed class TemplatePivotModelTest
{
    /// <summary>
    /// 验证既有透视新建请求仍使用旧默认行为。
    /// </summary>
    [Fact]
    public void ExistingPivotRequest_ShouldKeepDefaultTemplateModeDisabled()
    {
        var request = ExcelExport.Workbook(builder => builder
            .AddSheet("Data", new[] { new TemplatePivotBoundaryRow { Region = "North", Amount = 10 } })
            .PivotTable(Definition("SalesByRegion", "Data", "Summary")));

        Assert.False(request.PreserveTemplatePivotTables);
        Assert.Empty(request.PivotTableUpdates);
        Assert.Equal("SalesByRegion", Assert.Single(request.PivotTables).Name);
    }

    /// <summary>
    /// 验证 Request 不受创建后 Builder 的 Sheet、更新及保留模式变更影响。
    /// </summary>
    [Fact]
    public void Request_ShouldIsolateBuilderMutationsMadeAfterCreation()
    {
        var row = new TemplatePivotBoundaryRow { Region = "North", Amount = 10 };
        ExcelWorkbookExportBuilder? ordinaryBuilder = null;
        var ordinary = ExcelExport.Workbook(builder =>
        {
            ordinaryBuilder = builder;
            builder.AddSheet("Original", new[] { row });
        });

        ordinaryBuilder!
            .AddSheet("Later", new[] { row })
            .PreserveTemplatePivotTables();

        Assert.Equal(1, ordinary.SheetCount);
        Assert.Equal(new[] { "Original" }, ordinary.Sheets.Select(sheet => sheet.Name));
        Assert.False(ordinary.PreserveTemplatePivotTables);
        Assert.Empty(ordinary.PivotTableUpdates);
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelSheetExportRequest>)ordinary.Sheets).Clear());

        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        ExcelWorkbookExportBuilder? templateBuilder = null;
        var updated = ExcelExport.Workbook(builder =>
        {
            templateBuilder = builder;
            builder.UseTemplate(template, leaveOpen: true)
                .UpdatePivotTable(Definition("Initial", "TemplateData", "Summary"));
        });

        templateBuilder!
            .AddSheet("Later", new[] { row })
            .UpdatePivotTable(Definition("LaterUpdate", "TemplateData", "LaterSummary"));

        Assert.Equal(0, updated.SheetCount);
        Assert.Empty(updated.Sheets);
        Assert.True(updated.PreserveTemplatePivotTables);
        Assert.Equal(new[] { "Initial" }, updated.PivotTableUpdates.Select(item => item.Name));
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotTableDefinition>)updated.PivotTableUpdates).Clear());
    }

    /// <summary>
    /// 验证保留和更新模式都要求模板，并且模板格式只能是 XLSX。
    /// </summary>
    [Fact]
    public void TemplateModes_ShouldRequireXlsxTemplate()
    {
        Assert.Throws<InvalidOperationException>(() => ExcelExport.Workbook(builder =>
            builder.PreserveTemplatePivotTables()));
        Assert.Throws<InvalidOperationException>(() => ExcelExport.Workbook(builder =>
            builder.UpdatePivotTable(Definition("SalesByRegion", "Data", "Summary"))));

        using var preserveTemplate = new MemoryStream();
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder
            .Format(ExcelFormat.Xls)
            .UseTemplate(preserveTemplate, leaveOpen: true)
            .PreserveTemplatePivotTables()));
        using var updateTemplate = new MemoryStream();
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder
            .Format(ExcelFormat.Xls)
            .UseTemplate(updateTemplate, leaveOpen: true)
            .UpdatePivotTable(Definition("SalesByRegion", "Data", "Summary"))));
    }

    /// <summary>
    /// 验证仅保留模板透视时可以不添加新工作表。
    /// </summary>
    [Fact]
    public void PreserveMode_ShouldAllowTemplateOnlyRequest()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var request = ExcelExport.Workbook(builder => builder
            .UseTemplate(template, leaveOpen: true)
            .PreserveTemplatePivotTables());

        Assert.Equal(0, request.SheetCount);
        Assert.Same(template, request.Template);
        Assert.True(request.PreserveTemplatePivotTables);
        Assert.Empty(request.PivotTables);
        Assert.Empty(request.PivotTableUpdates);
    }

    /// <summary>
    /// 验证更新定义在 Builder 和 Request 阶段均保持深快照。
    /// </summary>
    [Fact]
    public void Update_ShouldSnapshotNestedDefinitionsAndExposeReadOnlyCollections()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var rowFields = new List<ExcelPivotAxisFieldDefinition>
        {
            new() { Field = "Region", SortOrder = ExcelPivotSortOrder.Descending }
        };
        var columnFields = new List<ExcelPivotAxisFieldDefinition>
        {
            new() { Field = "Month", SubtotalMode = ExcelPivotSubtotalMode.None }
        };
        var values = new List<ExcelPivotValueDefinition>
        {
            new() { Field = "Amount", Caption = "Revenue", NumberFormat = "0.00" }
        };
        var filterValues = new List<object> { "Open", 3, null };
        var filters = new List<ExcelPivotFilterDefinition>
        {
            new() { Field = "Status", Values = filterValues }
        };
        var range = new ExcelRangeDefinition
        {
            StartRow = 2, StartColumn = 1, EndRow = 20, EndColumn = 5
        };
        var definition = new ExcelPivotTableDefinition
        {
            Name = "SalesByRegion",
            SourceSheet = "TemplateData",
            SourceRange = range,
            TargetSheet = "Summary",
            TargetRow = 4,
            TargetColumn = 2,
            RowFields = rowFields,
            ColumnFields = columnFields,
            ValueFields = values,
            Filters = filters,
            MaxExpandedCells = 4096
        };

        var request = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(template, leaveOpen: true)
                .UpdatePivotTable(definition);
            rowFields[0] = new ExcelPivotAxisFieldDefinition { Field = "ChangedRow" };
            columnFields.Clear();
            values[0] = new ExcelPivotValueDefinition { Field = "ChangedValue", Caption = "Changed" };
            filterValues[0] = "Closed";
            filterValues.Add("Pending");
            filters[0] = new ExcelPivotFilterDefinition
            {
                Field = "ChangedFilter",
                Values = new object[] { "Changed" }
            };
        });

        var snapshot = Assert.Single(request.PivotTableUpdates);
        Assert.True(request.PreserveTemplatePivotTables);
        Assert.Equal("SalesByRegion", snapshot.Name);
        Assert.Equal("TemplateData", snapshot.SourceSheet);
        Assert.Equal("Summary", snapshot.TargetSheet);
        Assert.Equal(new[] { 2, 1, 20, 5 }, new[]
        {
            snapshot.SourceRange.StartRow, snapshot.SourceRange.StartColumn,
            snapshot.SourceRange.EndRow, snapshot.SourceRange.EndColumn
        });
        Assert.Equal(new[] { "Region" }, snapshot.RowFields.Select(field => field.Field));
        Assert.Equal(ExcelPivotSortOrder.Descending, snapshot.RowFields[0].SortOrder);
        Assert.Equal(new[] { "Month" }, snapshot.ColumnFields.Select(field => field.Field));
        Assert.Equal(ExcelPivotSubtotalMode.None, snapshot.ColumnFields[0].SubtotalMode);
        Assert.Equal("Revenue", Assert.Single(snapshot.ValueFields).Caption);
        Assert.Equal("0.00", snapshot.ValueFields[0].NumberFormat);
        Assert.Equal(new object[] { "Open", 3, null }, Assert.Single(snapshot.Filters).Values);
        Assert.Equal(4096, snapshot.MaxExpandedCells);

        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotTableDefinition>)request.PivotTableUpdates).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotAxisFieldDefinition>)snapshot.RowFields).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotAxisFieldDefinition>)snapshot.ColumnFields).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotValueDefinition>)snapshot.ValueFields).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelPivotFilterDefinition>)snapshot.Filters).Clear());
        Assert.Throws<NotSupportedException>(() =>
            ((IList<object>)snapshot.Filters[0].Values).Clear());
    }

    /// <summary>
    /// 验证大小写无关的重复更新被拒绝且 Builder 可继续配置。
    /// </summary>
    [Fact]
    public void DuplicateUpdates_ShouldRejectAndKeepBuilderReusable()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var request = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(template, leaveOpen: true)
                .UpdatePivotTable(Definition("SalesByRegion", "Data", "Summary"));
            Assert.Throws<ArgumentException>(() =>
                builder.UpdatePivotTable(Definition("salesbyregion", "Data", "Summary2")));
            builder.UpdatePivotTable(Definition("SalesByMonth", "Data", "Summary2"));
        });

        Assert.Equal(new[] { "SalesByRegion", "SalesByMonth" },
            request.PivotTableUpdates.Select(update => update.Name));
    }

    /// <summary>
    /// 验证新建定义与更新定义不能使用重复名称且配置失败后可复用。
    /// </summary>
    [Fact]
    public void NewAndUpdatedNames_ShouldRejectCaseInsensitiveCollisionsAndRemainReusable()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var updateAfterNew = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(template, leaveOpen: true)
                .AddSheet("Data", new[] { new TemplatePivotBoundaryRow { Region = "North", Amount = 10 } })
                .PivotTable(Definition("SalesByRegion", "Data", "NewSummary"));
            Assert.Throws<ArgumentException>(() =>
                builder.UpdatePivotTable(Definition("salesbyregion", "Data", "ExistingSummary")));
            builder.UpdatePivotTable(Definition("SalesByMonth", "TemplateData", "ExistingSummary"));
        });

        Assert.Single(updateAfterNew.PivotTables);
        Assert.Equal("SalesByMonth", Assert.Single(updateAfterNew.PivotTableUpdates).Name);

        using var secondTemplate = new MemoryStream(new byte[] { 1, 2, 3 });
        var newAfterUpdate = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(secondTemplate, leaveOpen: true)
                .AddSheet("Data", new[] { new TemplatePivotBoundaryRow { Region = "North", Amount = 10 } })
                .UpdatePivotTable(Definition("SalesByRegion", "TemplateData", "ExistingSummary"));
            Assert.Throws<ArgumentException>(() =>
                builder.PivotTable(Definition("salesbyregion", "Data", "NewSummary")));
            builder.PivotTable(Definition("SalesByMonth", "Data", "NewSummary"));
        });

        Assert.Equal("SalesByMonth", Assert.Single(newAfterUpdate.PivotTables).Name);
        Assert.Equal("SalesByRegion", Assert.Single(newAfterUpdate.PivotTableUpdates).Name);
    }

    /// <summary>
    /// 验证无效更新不会污染 Builder，随后可继续添加有效更新。
    /// </summary>
    [Fact]
    public void InvalidUpdate_ShouldNotPoisonBuilder()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var request = ExcelExport.Workbook(builder =>
        {
            builder.UseTemplate(template, leaveOpen: true);
            Assert.Throws<ArgumentException>(() => builder.UpdatePivotTable(new ExcelPivotTableDefinition
            {
                Name = "Invalid",
                SourceSheet = "TemplateData",
                TargetSheet = "Summary",
                RowField = "Region",
                ValueField = "Amount"
            }));
            builder.UpdatePivotTable(Definition("Valid", "TemplateData", "Summary"));
        });

        Assert.Equal("Valid", Assert.Single(request.PivotTableUpdates).Name);
    }

    /// <summary>
    /// 创建包含兼容旧字段的透视定义。
    /// </summary>
    /// <param name="name">透视表名称。</param>
    /// <param name="sourceSheet">模板或工作簿中的来源工作表名称。</param>
    /// <param name="targetSheet">模板或工作簿中的目标工作表名称。</param>
    /// <returns>有效的 Sum 透视定义。</returns>
    private static ExcelPivotTableDefinition Definition(string name, string sourceSheet, string targetSheet) =>
        new()
        {
            Name = name,
            SourceSheet = sourceSheet,
            SourceRange = new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1
            },
            TargetSheet = targetSheet,
            RowField = "Region",
            ValueField = "Amount"
        };
}
