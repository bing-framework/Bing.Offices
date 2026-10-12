using System.IO;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exports;
using Bing.Offices.Tests.PivotTemplates;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML 实际透视模板依赖。
/// </summary>
public sealed class ExcelPivotTemplateManifestTest
{
    /// <summary>
    /// 验证真实模板依赖、快照和异常边界。
    /// </summary>
    /// <param name="scenario">职责场景。</param>
    [Theory]
    [InlineData("normal")]
    [InlineData("shared")]
    [InlineData("renamed")]
    [InlineData("root-owner")]
    [InlineData("unknown-cache")]
    [InlineData("duplicate-name")]
    [InlineData("duplicate-cache")]
    [InlineData("duplicate-relationship")]
    [InlineData("orphan-pivot")]
    [InlineData("missing-records")]
    [InlineData("missing-cache-link")]
    [InlineData("external")]
    [InlineData("dtd")]
    [InlineData("canceled")]
    [InlineData("unknown-source")]
    [InlineData("invalid-range")]
    [InlineData("reversed-range")]
    [InlineData("overflow-range")]
    [InlineData("missing-range")]
    [InlineData("wrong-sheet-root")]
    [InlineData("strict")]
    [InlineData("wrong-worksheet-type")]
    [InlineData("duplicate-owner")]
    [InlineData("mismatched-cache-link")]
    public void Read_ShouldPreserveFullDependenciesAndRejectInvalidTemplate(string scenario)
    {
        ExcelPivotTemplateManifestTestCases.Verify(scenario, CreateTemplate(),
            new[] { "Category", "Amount", "OrderId" }, "A1:C3");
    }

    /// <summary>
    /// 验证资源限制和非标准工作簿位置。
    /// </summary>
    /// <param name="scenario">资源场景。</param>
    [Theory]
    [InlineData("input")]
    [InlineData("entries")]
    [InlineData("entry")]
    [InlineData("total")]
    [InlineData("ratio")]
    [InlineData("characters")]
    [InlineData("depth")]
    [InlineData("relationships")]
    [InlineData("relationship-depth")]
    [InlineData("renamed")]
    public void Read_WithResourceLimits_ShouldEnforceBudgetAndPreserveInput(string scenario) =>
        ExcelPivotTemplateManifestTestCases.VerifyResource(scenario, CreateTemplate(), "ClosedXML");

    /// <summary>
    /// 验证模板更新的精确定位和配置错误。
    /// </summary>
    /// <param name="scenario">预检场景。</param>
    [Theory]
    [InlineData("valid")]
    [InlineData("preserve-only")]
    [InlineData("new-name-conflict")]
    [InlineData("new-target-conflict")]
    [InlineData("unknown-name")]
    [InlineData("wrong-sheet")]
    [InlineData("unknown-source")]
    [InlineData("source-overflow")]
    [InlineData("row-overflow")]
    [InlineData("column-overflow")]
    [InlineData("canceled")]
    public void CreatePlan_ShouldLocateExactUpdateAndRejectInvalidConfiguration(string scenario) =>
        ExcelPivotTemplateManifestTestCases.VerifyPlan(scenario, CreateTemplate(),
            new[] { "Category", "Amount", "OrderId" });

    /// <summary>
    /// 验证 ClosedXML 模板快照的同步与异步复制、预检及清理边界。
    /// </summary>
    /// <param name="scenario">资源与失败场景。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData("success", false)]
    [InlineData("success", true)]
    [InlineData("budget", false)]
    [InlineData("budget", true)]
    [InlineData("unknown-update", false)]
    [InlineData("unknown-update", true)]
    [InlineData("malformed", false)]
    [InlineData("malformed", true)]
    [InlineData("pre-cancel", false)]
    [InlineData("pre-cancel", true)]
    [InlineData("copy-cancel", false)]
    [InlineData("copy-cancel", true)]
    [InlineData("copy-fail", false)]
    [InlineData("copy-fail", true)]
    public async Task Snapshot_ShouldVerifySyncAndAsyncCopyAndFailureBoundaries(string scenario,
        bool asynchronous)
    {
        await ExcelPivotTemplateManifestTestCases.VerifySnapshotAsync(scenario, asynchronous,
            CreateTemplate(), new[] { "Category", "Amount", "OrderId" }, "ClosedXML")
            .ConfigureAwait(false);
    }

    /// <summary>
    /// 创建包含两个透视的完整模板。
    /// </summary>
    /// <returns>通过公开导出生成的工作簿内容。</returns>
    private static byte[] CreateTemplate()
    {
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", new[]
        {
            new AdvancedPivotNumberFormatRow { Category = "North", Amount = 10, OrderId = "A" },
            new AdvancedPivotNumberFormatRow { Category = "South", Amount = 20, OrderId = "B" }
        }).PivotTable(Definition("ByGroupOne", "Summary"))
            .PivotTable(Definition("ByGroupTwo", "Summary2")));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建采购透视定义。
    /// </summary>
    /// <param name="name">透视名称。</param>
    /// <param name="target">目标工作表。</param>
    /// <returns>有界来源的基础透视。</returns>
    private static ExcelPivotTableDefinition Definition(string name, string target) => new()
    {
        Name = name, SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 2 },
        TargetSheet = target, TargetRow = 1, TargetColumn = 2,
        RowField = "Category", ValueField = "Amount"
    };
}
