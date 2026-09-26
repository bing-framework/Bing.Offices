using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML 名称列表原生校验的职责级直接测试。
/// </summary>
public sealed class NamedListValidationTest
{
    /// <summary>
    /// 验证局部名称优先于全局名称，并支持单格及含单引号的跨 Sheet 引用。
    /// </summary>
    [Fact]
    public void NamedList_ShouldResolveLocalGlobalSingleCellAndEscapedSheetNames()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        var escaped = workbook.Worksheets.Add("O'Brien Values");
        lookup.Cell("A1").Value = "Global";
        lookup.Cell("B1").Value = "RowA";
        lookup.Cell("C1").Value = "RowB";
        lookup.Cell("D1").Value = "RowC";
        data.Cell("C1").Value = "Local";
        escaped.Cell("A1").Value = "Escaped";
        workbook.DefinedNames.Add("Allowed", lookup.Range("A1:A1"));
        workbook.DefinedNames.Add("RowAllowed", lookup.Range("B1:D1"));
        data.DefinedNames.Add("Allowed", data.Range("C1:C1"));
        workbook.DefinedNames.Add("EscapedAllowed", escaped.Range("A1:A1"));

        Assert.True(Evaluate(data, "A2", "Allowed", "Local").IsValid);
        var shadowed = Evaluate(data, "A3", "Allowed", "Global");
        Assert.False(shadowed.IsValid);
        Assert.False(shadowed.IsUnsupported);
        Assert.True(Evaluate(data, "A4", "EscapedAllowed", "Escaped").IsValid);
        Assert.True(Evaluate(data, "A5", "RowAllowed", "RowB").IsValid);
    }

    /// <summary>
    /// 验证名称仅接受单个绝对有界的一维引用。
    /// </summary>
    /// <param name="refersTo">命名范围的引用表达式。</param>
    [Theory]
    [InlineData("OFFSET('Lookup'!$A$1,0,0,2,1)")]
    [InlineData("OtherName")]
    [InlineData("'Lookup'!$A$1:$A$2,'Lookup'!$C$1:$C$2")]
    [InlineData("'Lookup'!$A$1:$B$2")]
    [InlineData("'Lookup'!A1:A2")]
    [InlineData("'Lookup'!$A:$A")]
    public void NamedList_UnsupportedDefinitions_ShouldReturnStructuredUnsupported(string refersTo)
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        workbook.Worksheets.Add("Lookup");
        workbook.DefinedNames.Add("BadList", refersTo);

        var result = Evaluate(data, "A2", "BadList", "A");

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
        Assert.Equal("Workbook Data Validation 规则类型或公式暂不支持。", result.Message);
    }

    /// <summary>
    /// 验证失效局部名称遮蔽同名全局名称，不静默回退。
    /// </summary>
    [Fact]
    public void NamedList_InvalidLocalName_ShouldNotFallBackToGlobalName()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "Global";
        workbook.DefinedNames.Add("Allowed", lookup.Range("A1:A1"));
        data.DefinedNames.Add("Allowed", "OtherName");

        var result = Evaluate(data, "A2", "Allowed", "Global");

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
    }

    /// <summary>
    /// 验证循环别名和缺失 Sheet 名称保持结构化 Unsupported。
    /// </summary>
    /// <param name="cyclicAlias">是否使用循环别名；否则使用缺失工作表引用。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedList_CyclicAliasAndMissingSheet_ShouldRemainUnsupported(bool cyclicAlias)
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        if (cyclicAlias)
        {
            workbook.DefinedNames.Add("FirstList", "SecondList");
            workbook.DefinedNames.Add("SecondList", "FirstList");
        }
        else
            workbook.DefinedNames.Add("FirstList", "'Missing Sheet'!$A$1:$A$2");

        var result = Evaluate(data, "A2", "FirstList", "A");

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
    }

    /// <summary>
    /// 验证名称范围完整扫描，后续公式单元格不会被较早匹配掩盖。
    /// </summary>
    [Fact]
    public void NamedList_FormulaCellAfterMatch_ShouldRemainUnsupportedWithoutRecalculation()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "Match";
        lookup.Cell("A2").FormulaA1 = "\"Formula\"";
        workbook.DefinedNames.Add("Allowed", lookup.Range("A1:A2"));

        var result = Evaluate(data, "A2", "Allowed", "Match");

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
    }

    /// <summary>
    /// 验证旧显式单项列表保留空格，而逗号列表继续修剪并忽略空项。
    /// </summary>
    [Fact]
    public void ExplicitList_ShouldPreserveExistingSingleAndCommaSeparatedSemantics()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");

        Assert.True(Evaluate(data, "A2", "\" A \"", " A ").IsValid);
        Assert.False(Evaluate(data, "A3", "\" A \"", "A").IsValid);
        Assert.True(Evaluate(data, "A4", "\" A, ,B \"", "B").IsValid);
    }

    /// <summary>
    /// 验证空值策略和列表扫描入口的预取消检查。
    /// </summary>
    [Fact]
    public void NamedList_BlankAndCancellation_ShouldUseExistingPolicies()
    {
        using var workbook = new XLWorkbook();
        var data = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "A";
        lookup.Cell("A2").Value = "B";
        workbook.DefinedNames.Add("Allowed", lookup.Range("A1:A2"));

        Assert.True(Evaluate(data, "A2", "Allowed", string.Empty, ignoreBlanks: true).IsValid);
        var blank = Evaluate(data, "A3", "Allowed", string.Empty, ignoreBlanks: false);
        Assert.False(blank.IsValid);
        Assert.False(blank.IsUnsupported);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() =>
            Evaluate(data, "A4", "Allowed", "B", cancellationToken: cancellation.Token));
    }

    /// <summary>
    /// 验证保存重开后的名称列表通过同步和异步公共导入入口并保留源流所有权。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_NamedListValidValue_ShouldRoundTripAndKeepSourceOpen(bool asynchronous)
    {
        using var source = CreateWorkbook("Allowed", "Allowed", useFormulaCell: false);
        var importer = new ClosedXmlExcelImporter();
        var request = CreateRequest(ExcelUnsupportedFeaturePolicy.Fail);

        var result = asynchronous
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Allowed", Assert.Single(result.Workbook.Rows).Name);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证真实非法列表值即使在 Report 策略下也拒绝数据行。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_NamedListInvalidValue_ShouldRejectRowUnderReportPolicy(bool asynchronous)
    {
        using var source = CreateWorkbook("Rejected", "Allowed", useFormulaCell: false);

        var importer = new ClosedXmlExcelImporter();
        var result = asynchronous
            ? await importer.ImportAsync(source, CreateRequest(ExcelUnsupportedFeaturePolicy.Report))
            : importer.Import(source, CreateRequest(ExcelUnsupportedFeaturePolicy.Report));

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        AssertImportError(Assert.Single(result.Errors), "不符合 Workbook 列表校验。");
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证名称公式边界在 Report 下保留行、Fail 下拒绝行。
    /// </summary>
    /// <param name="policy">待验证的冲突或不支持功能处理策略。</param>
    /// <param name="expectedRows">预期保留的数据行数。</param>
    /// <param name="asynchronous">是否使用异步导入入口。</param>
    [Theory]
    [InlineData(ExcelUnsupportedFeaturePolicy.Report, 1, false)]
    [InlineData(ExcelUnsupportedFeaturePolicy.Report, 1, true)]
    [InlineData(ExcelUnsupportedFeaturePolicy.Fail, 0, false)]
    [InlineData(ExcelUnsupportedFeaturePolicy.Fail, 0, true)]
    public async Task Import_NamedListFormulaSource_ShouldRespectUnsupportedPolicy(
        ExcelUnsupportedFeaturePolicy policy, int expectedRows, bool asynchronous)
    {
        using var source = CreateWorkbook("Allowed", "Allowed", useFormulaCell: true);

        var importer = new ClosedXmlExcelImporter();
        var result = asynchronous
            ? await importer.ImportAsync(source, CreateRequest(policy))
            : importer.Import(source, CreateRequest(policy));

        Assert.False(result.IsSuccess);
        Assert.Equal(expectedRows, result.Workbook.Rows.Count);
        AssertImportError(Assert.Single(result.Errors),
            "Workbook Data Validation 规则类型或公式暂不支持。");
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步预取消不关闭调用方输入流。
    /// </summary>
    [Fact]
    public async Task ImportAsync_PreCanceled_ShouldKeepSourceOpen()
    {
        using var source = CreateWorkbook("Allowed", "Allowed", useFormulaCell: false);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().ImportAsync(source,
                CreateRequest(ExcelUnsupportedFeaturePolicy.Fail), cancellation.Token));

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 执行指定单元格的列表校验。
    /// </summary>
    /// <param name="worksheet">承载校验规则的工作表。</param>
    /// <param name="targetAddress">待校验单元格的 A1 地址。</param>
    /// <param name="expression">待验证的原生引用表达式。</param>
    /// <param name="text">待校验的单元格文本。</param>
    /// <param name="ignoreBlanks">是否允许空白值。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>列表规则的校验结果。</returns>
    private static ClosedXmlWorkbookValidationPipeline.Result Evaluate(IXLWorksheet worksheet,
        string targetAddress, string expression, string text, bool ignoreBlanks = false,
        CancellationToken cancellationToken = default)
    {
        var validation = worksheet.Range(targetAddress).CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.List;
        validation.Value = expression;
        validation.IgnoreBlanks = ignoreBlanks;
        return ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell(targetAddress), text,
            text, CultureInfo.InvariantCulture, false, cancellationToken);
    }

    /// <summary>
    /// 创建包含命名列表校验的测试工作簿。
    /// </summary>
    /// <param name="importedValue">待导入的单元格值。</param>
    /// <param name="allowedValue">命名列表允许的值。</param>
    /// <param name="useFormulaCell">是否在命名列表中包含公式单元格。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的内存流。</returns>
    private static MemoryStream CreateWorkbook(string importedValue, string allowedValue,
        bool useFormulaCell)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var data = workbook.Worksheets.Add("Data");
            var lookup = workbook.Worksheets.Add("Lookup Values");
            data.Cell("A1").Value = nameof(Row.Name);
            data.Cell("A2").Value = importedValue;
            lookup.Cell("A1").Value = allowedValue;
            if (useFormulaCell)
                lookup.Cell("A2").FormulaA1 = "\"Formula\"";
            else
                lookup.Cell("A2").Value = "Other";
            workbook.DefinedNames.Add("AllowedValues", lookup.Range("A1:A2"));
            var validation = data.Range("A2").CreateDataValidation();
            validation.AllowedValues = XLAllowedValues.List;
            validation.Value = "AllowedValues";
            validation.IgnoreBlanks = false;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建启用工作簿规则校验的导入请求。
    /// </summary>
    /// <param name="policy">待验证的冲突或不支持功能处理策略。</param>
    /// <returns>配置了指定不支持功能策略的导入请求。</returns>
    private static ExcelWorkbookImportRequest<Model> CreateRequest(ExcelUnsupportedFeaturePolicy policy) =>
        ExcelImport.Workbook<Model>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .UnsupportedFeaturePolicy(policy)
            .Sheet<Row>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 验证工作簿校验错误的位置和消息。
    /// </summary>
    /// <param name="error">待验证的结构化导入错误。</param>
    /// <param name="message">预期错误消息。</param>
    private static void AssertImportError(ExcelImportError error, string message)
    {
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(Row.Name), error.PropertyName);
        Assert.Equal(message, error.Message);
    }

    /// <summary>
    /// 承载命名列表导入结果的测试工作簿。
    /// </summary>
    private sealed class Model
    {
        /// <summary>
        /// 获取导入结果行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 用于导入或导出测试的数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置测试行名称。
        /// </summary>
        public string Name { get; set; }
    }
}
