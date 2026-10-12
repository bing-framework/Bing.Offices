using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 实体失败工作簿的资源预算和暂存清理边界。
/// </summary>
public sealed class EntityFailureWorkbookResourceTest
{
    /// <summary>
    /// 验证实体错误候选行预算在目标写出前生效。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CandidateErrorRowBudget_ShouldRejectBeforeDestinationWrite(bool xls)
    {
        using var input = new MemoryStream(CreateSimpleFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var result = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().ImportEntity(input, CreateSimpleLayout(),
                CreateFailureOptions(failure, maxCandidateErrorRows: 1)));

        Assert.Equal(BingOfficesStage.Preflight, result.Stage);
        Assert.Equal(BingOfficesOperation.Import, result.Operation);
        Assert.Equal(0, failure.Length);
        Assert.True(input.CanRead);
    }

    /// <summary>
    /// 验证动态字典结构的复制单元格预算在实体写出前生效。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DynamicDictionary_CopiedCellBudget_ShouldRejectBeforeDestinationWrite(bool xls)
    {
        using var input = new MemoryStream(CreateDynamicFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var result = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().ImportEntity(input, CreateDynamicLayout(),
                CreateFailureOptions(failure, maxCopiedCells: 3)));

        Assert.Equal(BingOfficesStage.Preflight, result.Stage);
        Assert.Equal(BingOfficesOperation.Import, result.Operation);
        Assert.Equal(0, failure.Length);
        Assert.True(input.CanRead);
    }

    /// <summary>
    /// 验证父子列表结构的目标对象预算在实体写出前生效。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RelatedLists_EstimatedObjectBudget_ShouldRejectBeforeDestinationWrite(bool xls)
    {
        using var input = new MemoryStream(CreateRelatedFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var result = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().ImportEntity(input, CreateRelationLayout(),
                CreateFailureOptions(failure, maxEstimatedTargetObjects: 11)));

        Assert.Equal(BingOfficesStage.Preflight, result.Stage);
        Assert.Equal(BingOfficesOperation.Import, result.Operation);
        Assert.Equal(0, failure.Length);
        Assert.True(input.CanRead);
    }

    /// <summary>
    /// 验证实体失败工作簿序列化超限时保留目标文件并清理暂存文件。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SerializedByteBudget_ShouldPreserveTargetAndCleanupTemporaryFile(bool xls)
    {
        using var directory = new TestDirectory();
        var path = Path.Combine(directory.Path, xls ? "failure.xls" : "failure.xlsx");
        var sentinel = new byte[] { 4, 8, 15, 16, 23, 42 };
        File.WriteAllBytes(path, sentinel);
        using var input = new MemoryStream(CreateSimpleFailureWorkbook(xls), writable: false);
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path,
                MaxSerializedBytes = 1
            });

        var result = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new NpoiExcelImporter().ImportEntity(input, CreateSimpleLayout(), options));

        Assert.Equal(BingOfficesStage.Serialize, result.Stage);
        Assert.Equal(BingOfficesOperation.Import, result.Operation);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 创建实体失败工作簿选项。
    /// </summary>
    /// <param name="destination">接收失败工作簿的目标流。</param>
    /// <param name="maxCandidateErrorRows">候选错误行预算。</param>
    /// <param name="maxCopiedCells">复制单元格预算。</param>
    /// <param name="maxEstimatedTargetObjects">目标对象预算。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions CreateFailureOptions(Stream destination,
        int? maxCandidateErrorRows = null, long? maxCopiedCells = null,
        long? maxEstimatedTargetObjects = null) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination,
                MaxCandidateErrorRows = maxCandidateErrorRows,
                MaxCopiedCells = maxCopiedCells,
                MaxEstimatedTargetObjects = maxEstimatedTargetObjects
            });

    /// <summary>
    /// 创建隐式动态字典布局。
    /// </summary>
    /// <returns>动态字典实体布局。</returns>
    private static ExcelEntityLayout<DynamicDocument> CreateDynamicLayout()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(DynamicLine.Code), Title = "Code" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "score", Title = "Score", DataTypeName = "int32",
                    Order = 0, ColumnIndex = 1 }
            }
        };
        return ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("动态", "A1", document => document.Lines,
                region => region.Mapping(mapping)));
    }

    /// <summary>
    /// 创建包含父子列表关系的实体布局。
    /// </summary>
    /// <returns>父子列表实体布局。</returns>
    private static ExcelEntityLayout<RelationDocument> CreateRelationLayout() =>
        ExcelEntity.Layout<RelationDocument>(builder => builder
            .ListRegion("父项", "A1", document => document.Parents)
            .ListRegion("子项", "A1", document => document.Children)
            .HasMany(document => document.Parents, document => document.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 创建单列表实体布局。
    /// </summary>
    /// <returns>单列表实体布局。</returns>
    private static ExcelEntityLayout<SimpleDocument> CreateSimpleLayout() =>
        ExcelEntity.Layout<SimpleDocument>(builder =>
            builder.ListRegion("数据", "A1", document => document.Lines));

    /// <summary>
    /// 创建动态列转换失败的源工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 格式。</param>
    /// <returns>源工作簿字节。</returns>
    private static byte[] CreateDynamicFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("动态");
        WriteHeader(sheet, "Code", "Score");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("C-1");
        row.CreateCell(1).SetCellValue("wrong");
        return WriteWorkbook(workbook);
    }

    /// <summary>
    /// 创建子列表转换失败的源工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 格式。</param>
    /// <returns>源工作簿字节。</returns>
    private static byte[] CreateRelatedFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        WriteHeader(parents, "Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        var children = workbook.CreateSheet("子项");
        WriteHeader(children, "ParentId", "Name", "Quantity");
        var row = children.CreateRow(1);
        row.CreateCell(0).SetCellValue("P-1");
        row.CreateCell(1).SetCellValue("detail");
        row.CreateCell(2).SetCellValue("wrong");
        return WriteWorkbook(workbook);
    }

    /// <summary>
    /// 创建单列表必填字段失败的源工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 格式。</param>
    /// <returns>源工作簿字节。</returns>
    private static byte[] CreateSimpleFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("数据");
        WriteHeader(sheet, "Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("wrong");
        sheet.CreateRow(2).CreateCell(0).SetCellValue("also-wrong");
        return WriteWorkbook(workbook);
    }

    /// <summary>
    /// 写入工作表表头。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="headers">表头文本。</param>
    private static void WriteHeader(ISheet sheet, params string[] headers)
    {
        var row = sheet.CreateRow(0);
        for (var index = 0; index < headers.Length; index++)
            row.CreateCell(index).SetCellValue(headers[index]);
    }

    /// <summary>
    /// 将 NPOI 工作簿序列化为字节。
    /// </summary>
    /// <param name="workbook">待序列化的工作簿。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] WriteWorkbook(IWorkbook workbook)
    {
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 保存动态字典实体根对象。
    /// </summary>
    private sealed class DynamicDocument
    {
        /// <summary>
        /// 获取动态明细集合。
        /// </summary>
        public List<DynamicLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存动态字典实体行。
    /// </summary>
    private sealed class DynamicLine
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
    /// 保存父子列表实体根对象。
    /// </summary>
    private sealed class RelationDocument
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
    /// 保存父列表行。
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
    /// 保存子列表行。
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
    /// 保存简单列表实体根对象。
    /// </summary>
    private sealed class SimpleDocument
    {
        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<SimpleLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存简单列表实体行。
    /// </summary>
    private sealed class SimpleLine
    {
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 为单项测试创建并清理临时目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取测试专用目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "npoi-entity-resource-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, true);
    }
}
