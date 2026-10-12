using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML Entity 失败工作簿资源上限拒绝时不触碰目标文件。
/// </summary>
public sealed class EntityFailureWorkbookResourceTest
{
    /// <summary>
    /// 验证动态字典的复制单元格预算在实体写出前生效。
    /// </summary>
    [Fact]
    public void DynamicDictionary_CopiedCellBudget_ShouldRejectBeforeDestinationWrite()
    {
        using var directory = new TestDirectory();
        using var source = CreateDynamicFailureWorkbook();
        var path = Path.Combine(directory.Path, "dynamic.xlsx");
        var sentinel = new byte[] { 1, 3, 5, 7 };
        File.WriteAllBytes(path, sentinel);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(source, CreateDynamicLayout(),
                FailureOptions(directory.Path, path, maxCopiedCells: 3)));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证父子列表的目标对象预算在实体写出前生效。
    /// </summary>
    [Fact]
    public void RelatedLists_EstimatedObjectBudget_ShouldRejectBeforeDestinationWrite()
    {
        using var directory = new TestDirectory();
        using var source = CreateRelatedFailureWorkbook();
        var path = Path.Combine(directory.Path, "related.xlsx");
        var sentinel = new byte[] { 2, 4, 6, 8 };
        File.WriteAllBytes(path, sentinel);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(source, CreateRelationLayout(),
                FailureOptions(directory.Path, path, maxEstimatedTargetObjects: 11)));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证候选错误行超限时保留目标文件且不留下暂存文件。
    /// </summary>
    [Fact]
    public void MaxCandidateErrorRows_ShouldPreserveTargetAndCleanup()
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 8, 6, 4, 2 };
        File.WriteAllBytes(path, sentinel);
        var options = FailureOptions(directory.Path, path, maxCandidateErrorRows: 1);
        var importer = new ClosedXmlExcelImporter();

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntity(source, CreateLayout(), options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证序列化字节超限时保留目标文件且清理失败工作簿暂存文件。
    /// </summary>
    [Fact]
    public void MaxSerializedBytes_ShouldPreserveTargetAndCleanup()
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook();
        var path = Path.Combine(directory.Path, "failure.xlsx");
        var sentinel = new byte[] { 3, 1, 4, 1, 5 };
        File.WriteAllBytes(path, sentinel);
        var options = FailureOptions(directory.Path, path, maxSerializedBytes: 1);
        var importer = new ClosedXmlExcelImporter();

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntity(source, CreateLayout(), options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesStage.Serialize, exception.Stage);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建把必填明细映射到 Data 工作表的布局。
    /// </summary>
    /// <returns>Data 工作表实体布局。</returns>
    private static ExcelEntityLayout<Root> CreateLayout() =>
        ExcelEntity.Layout<Root>(builder => builder.ListRegion("Data", "A1", root => root.Rows));

    /// <summary>
    /// 创建动态字典实体布局。
    /// </summary>
    /// <returns>动态字典实体布局。</returns>
    private static ExcelEntityLayout<DynamicRoot> CreateDynamicLayout()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(DynamicRow.Code), Title = "Code" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "score", Title = "Score", DataTypeName = "int32",
                    Order = 0, ColumnIndex = 1 }
            }
        };
        return ExcelEntity.Layout<DynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Rows,
                region => region.Mapping(mapping)));
    }

    /// <summary>
    /// 创建父子列表实体布局。
    /// </summary>
    /// <returns>父子列表实体布局。</returns>
    private static ExcelEntityLayout<RelationRoot> CreateRelationLayout() =>
        ExcelEntity.Layout<RelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 创建包含两条必填字段错误的真实 XLSX 流。
    /// </summary>
    /// <returns>位于起始位置的工作簿流。</returns>
    private static MemoryStream CreateInvalidWorkbook()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("B1").Value = "Quantity";
            sheet.Cell("A2").Value = string.Empty;
            sheet.Cell("B2").Value = 1;
            sheet.Cell("A3").Value = string.Empty;
            sheet.Cell("B3").Value = 2;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建动态列转换失败的真实 XLSX 流。
    /// </summary>
    /// <returns>位于起始位置的工作簿流。</returns>
    private static MemoryStream CreateDynamicFailureWorkbook()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Dynamic");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("B1").Value = "Score";
            sheet.Cell("A2").Value = "C-1";
            sheet.Cell("B2").Value = "wrong";
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建子列表转换失败的真实 XLSX 流。
    /// </summary>
    /// <returns>位于起始位置的工作簿流。</returns>
    private static MemoryStream CreateRelatedFailureWorkbook()
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var parents = workbook.Worksheets.Add("Parents");
            parents.Cell("A1").Value = "Id";
            parents.Cell("A2").Value = "P-1";
            var children = workbook.Worksheets.Add("Children");
            children.Cell("A1").Value = "ParentId";
            children.Cell("B1").Value = "Name";
            children.Cell("C1").Value = "Quantity";
            children.Cell("A2").Value = "P-1";
            children.Cell("B2").Value = "detail";
            children.Cell("C2").Value = "wrong";
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 使用指定文件目标和资源限制创建实体导入选项。
    /// </summary>
    /// <param name="temporaryDirectory">失败工作簿暂存目录。</param>
    /// <param name="path">失败工作簿目标文件。</param>
    /// <param name="maxSerializedBytes">失败工作簿最大序列化字节数。</param>
    /// <param name="maxCandidateErrorRows">最大候选错误行数。</param>
    /// <returns>用于继续采集错误并生成失败工作簿的实体导入选项。</returns>
    private static ExcelEntityImportOptions FailureOptions(string temporaryDirectory, string path,
        long? maxSerializedBytes = null, int? maxCandidateErrorRows = null,
        long? maxCopiedCells = null, long? maxEstimatedTargetObjects = null) =>
        new(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = temporaryDirectory,
                MaxSerializedBytes = maxSerializedBytes,
                MaxCandidateErrorRows = maxCandidateErrorRows,
                MaxCopiedCells = maxCopiedCells,
                MaxEstimatedTargetObjects = maxEstimatedTargetObjects
            });

    /// <summary>
    /// 为资源限制测试创建和清理独立暂存目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取当前测试专用目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "closedxml-entity-failure-resource-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }

    /// <summary>
    /// 保存实体失败工作簿测试中的根对象。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取明细行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 保存动态列测试根对象。
    /// </summary>
    private sealed class DynamicRoot
    {
        /// <summary>
        /// 获取动态明细集合。
        /// </summary>
        public List<DynamicRow> Rows { get; } = new();
    }

    /// <summary>
    /// 保存动态列测试行。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置业务编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取动态字段字典。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Fields { get; set; } = new();
    }

    /// <summary>
    /// 保存父子列表测试根对象。
    /// </summary>
    private sealed class RelationRoot
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存父列表测试行。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关联子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存子列表测试行。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 保存实体失败工作簿测试中的必填行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置必填编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置用于保留明细行的数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
