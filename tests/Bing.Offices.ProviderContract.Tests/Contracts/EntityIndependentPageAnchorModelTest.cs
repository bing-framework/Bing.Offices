using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Entities;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证实体独立逐页页眉与尾部命名锚点的模型合同。
/// </summary>
public sealed class EntityIndependentPageAnchorModelTest
{
    /// <summary>
    /// 验证公开构建入口保存逐页描述并在解析起点副本中保持快照。
    /// </summary>
    [Fact]
    public void PageAnchorNamed_ShouldSnapshotAndResolveStart()
    {
        var layout = CreateNamedLayout();
        var region = layout.ListRegions.Single();

        Assert.Equal("HeaderPage", region.PageHeaderAnchorPrefix);
        Assert.Equal(ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete,
            region.PageHeaderAnchorStaleNamePolicy);
        Assert.Equal("FooterPage", region.PageFooterAnchorPrefix);
        Assert.Equal(ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete,
            region.PageFooterAnchorStaleNamePolicy);
        Assert.Equal("HEADER", region.PageHeader.MarkerText);
        Assert.Equal("PAGE_FOOTER", region.PageFooter.MarkerText);
        Assert.Equal(2, region.PageBreakRows);

        var resolved = region.WithResolvedStart(ExcelEntityCellReference.Parse("B5"));

        Assert.NotSame(region, resolved);
        Assert.Equal("B5", resolved.Start.Address);
        Assert.Equal(region.PageHeaderAnchorPrefix, resolved.PageHeaderAnchorPrefix);
        Assert.Equal(region.PageFooterAnchorPrefix, resolved.PageFooterAnchorPrefix);
        Assert.Equal(region.PageHeaderAnchorStaleNamePolicy,
            resolved.PageHeaderAnchorStaleNamePolicy);
        Assert.Equal(region.PageFooterAnchorStaleNamePolicy,
            resolved.PageFooterAnchorStaleNamePolicy);
        Assert.Same(region.PageFooter, resolved.PageFooter);
    }

    /// <summary>
    /// 验证配置委托失败或策略无效后仍可复用构建器。
    /// </summary>
    [Fact]
    public void PageAnchorConfigurationFailure_ShouldLeaveBuilderReusable()
    {
        var builder = new ExcelEntityListRegionBuilder<AnchorLine>();
        builder.PageBreak(2);

        Assert.Throws<ArgumentException>(() => builder.PageFooter("PAGE",
            footer => footer.FormulaSumDetailRowsAbove("A1", "B")));

        Assert.Throws<ArgumentOutOfRangeException>(() => builder.PageFooterNamed("FooterPage",
            "PAGE", staleNames: (ExcelEntityPageAnchorStaleNamePolicy)99));
        Assert.Throws<ArgumentException>(() => builder.PageFooterNamed("Bad Name", "PAGE"));

        builder.PageFooterNamed("FooterPage", "PAGE");

        var headerBuilder = new ExcelEntityListRegionBuilder<AnchorLine>();
        headerBuilder.PageBreak(2);
        Assert.Throws<InvalidOperationException>(() => headerBuilder.PageHeaderNamed(
            "HeaderPage", "A1", "A1", "HEADER",
            _ => throw new InvalidOperationException("delegate failure")));
        headerBuilder.PageHeaderNamed("HeaderPage", "A1", "A1", "HEADER");
    }

    /// <summary>
    /// 验证每页入口在缺少分页策略时拒绝配置，并支持先分页后配置。
    /// </summary>
    [Fact]
    public void PageFooter_ShouldRequirePageBreak()
    {
        var builder = new ExcelEntityListRegionBuilder<AnchorLine>();

        Assert.Throws<InvalidOperationException>(() => builder.PageFooter("PAGE"));
        builder.PageBreak(1).PageFooter("PAGE");

        var layout = ExcelEntity.Layout<AnchorDocument>(layoutBuilder => layoutBuilder
            .ListRegion("Data", "A1", document => document.Lines, region => region
                .PageBreak(1).PageFooter("PAGE")));
        Assert.Equal("PAGE", layout.ListRegions.Single().PageFooter.MarkerText);
    }

    /// <summary>
    /// 验证逐页系列、静态锚点和 marker 冲突均在布局构建阶段拒绝。
    /// </summary>
    [Fact]
    public void PageAnchorLayout_ShouldRejectNameAndMarkerConflicts()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .CellNamed("Data", "FooterPage_001", document => document.Title)
            .ListRegion("Data", "A3", document => document.Lines, region => region
                .Header(false).PageBreak(2).PageFooterNamed("FooterPage", "PAGE"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A3", document => document.Lines, region => region
                .Header(false).PageBreak(2)
                .PageHeaderNamed("Pages", "A1:B2", "A2", "PAGE")
                .PageFooter("PAGE"))));
    }

    /// <summary>
    /// 验证最终连续求和不能跨越每页尾部行段。
    /// </summary>
    [Fact]
    public void FinalContiguousSum_ShouldRejectPageFooterBoundary()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A1", document => document.Lines, region => region
                .PageBreak(2).PageFooter("PAGE")
                .Footer("TOTAL", footer => footer.FormulaSumContiguousRowsAbove("B1", "B")))));
    }

    /// <summary>
    /// 验证每页尾部连续求和不能跨越分页或分组小计。
    /// </summary>
    [Fact]
    public void PageFooterContiguousSum_ShouldRejectSubtotalBoundaries()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageFooter("PAGE", footer => footer.FormulaSumContiguousRowsAbove("B1", "B"))
                .PageSubtotal("SUB"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageFooter("PAGE", footer => footer.FormulaSumContiguousRowsAbove("B1", "B"))
                .GroupSubtotal(line => line.Name, "GROUP"))));
    }

    /// <summary>
    /// 验证分组小计连续求和不能跨越每页尾部。
    /// </summary>
    [Fact]
    public void GroupSubtotalContiguousSum_ShouldRejectPageFooterBoundary()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .GroupSubtotal(line => line.Name, "GROUP",
                    footer => footer.FormulaSumContiguousRowsAbove("B1", "B"))
                .PageFooter("PAGE"))));
    }

    /// <summary>
    /// 验证没有新配置时旧页眉与尾部调用的默认值保持兼容。
    /// </summary>
    [Fact]
    public void ExistingPageHeaderAndFooter_ShouldKeepDefaults()
    {
        var layout = ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegion("Data", "A3", document => document.Lines, region => region
                .Header(false).PageBreak(2)
                .PageHeader("A1:B2", "A2", "HEADER")
                .Footer("TOTAL")));
        var region = layout.ListRegions.Single();

        Assert.NotNull(region.PageHeader);
        Assert.Null(region.PageHeaderAnchorPrefix);
        Assert.Equal(ExcelEntityPageAnchorStaleNamePolicy.Reject,
            region.PageHeaderAnchorStaleNamePolicy);
        Assert.Null(region.PageFooter);
        Assert.Null(region.PageFooterAnchorPrefix);
        Assert.Equal(ExcelEntityPageAnchorStaleNamePolicy.Reject,
            region.PageFooterAnchorStaleNamePolicy);
    }

    /// <summary>
    /// 验证逐页名称策略的枚举值固定为公开合同规定的数值。
    /// </summary>
    [Fact]
    public void PageAnchorStaleNamePolicy_ShouldKeepContractValues()
    {
        Assert.Equal(0, (int)ExcelEntityPageAnchorStaleNamePolicy.Reject);
        Assert.Equal(1, (int)ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete);
    }

    /// <summary>
    /// 创建包含逐页页眉、每页尾部及命名起点的模型布局。
    /// </summary>
    /// <returns>已完成校验的实体布局。</returns>
    private static ExcelEntityLayout<AnchorDocument> CreateNamedLayout() =>
        ExcelEntity.Layout<AnchorDocument>(builder => builder
            .ListRegionNamed("Data", "LinesStart", document => document.Lines, region => region
                .Header(false).PageBreak(2)
                .PageHeaderNamed("HeaderPage", "A1:B2", "A2", "HEADER",
                    header => header.Cell("A1", "采购单"),
                    ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete)
                .PageFooterNamed("FooterPage", "PAGE_FOOTER",
                    footer => footer.Cell("B1", "签字"),
                    ExcelEntityPageAnchorStaleNamePolicy.RemoveObsolete)));

    /// <summary>
    /// 表示模型职责测试中的采购单聚合。
    /// </summary>
    private sealed class AnchorDocument
    {
        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置采购明细。
        /// </summary>
        public List<AnchorLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示模型职责测试中的采购明细。
    /// </summary>
    private sealed class AnchorLine
    {
        /// <summary>
        /// 获取或设置明细名称。
        /// </summary>
        public string Name { get; set; } = string.Empty;
    }
}
