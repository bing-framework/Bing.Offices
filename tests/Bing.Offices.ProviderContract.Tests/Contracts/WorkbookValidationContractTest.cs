using System.IO;
using System;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Models;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 Workbook 原生 Data Validation 的规则矩阵入口。
/// </summary>
public sealed class WorkbookValidationContractTest
{
    /// <summary>
    /// ClosedXML 和 NPOI 应执行同一列表规则。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task ExplicitListRule_ShouldReturnWorkbookValidationError(string provider)
    {
        using var source = Golden("validation/list-invalid.xlsx");
        var request = CreateRequest();
        var importer = ProviderDrivers.Get(provider).CreateImporter();
        var result = await importer.ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(nameof(ValidationContractRow.Code), error.PropertyName);
    }

    /// <summary>
    /// 数值比较规则应按 ClosedXML 与 NPOI 的同一边界处理。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task WholeNumberBetweenRule_ShouldAcceptValue(string provider)
    {
        using var source = Golden("validation/whole-between-valid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
        Assert.Equal(5, result.Workbook.Rows[0].Quantity);
    }

    /// <summary>
    /// 数值比较失败应返回公共 WorkbookValidation 错误。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task WholeNumberLessThanRule_ShouldRejectValue(string provider)
    {
        using var source = Golden("validation/whole-less-than-invalid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, Assert.Single(result.Errors).Code);
    }

    /// <summary>
    /// 不支持的原生公式在 Report 模式应保留行，在 Fail 模式应拒绝行。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task UnsupportedRule_ShouldRespectReportAndFailPolicy(string provider)
    {
        var reportRequest = CreateRequest(ExcelUnsupportedFeaturePolicy.Report);
        var failRequest = CreateRequest(ExcelUnsupportedFeaturePolicy.Fail);

        using var reportSource = Golden("validation/custom-unsupported.xlsx");
        var reportResult = await ProviderDrivers.Get(provider).CreateImporter()
            .ImportAsync(reportSource, reportRequest);
        using var failSource = Golden("validation/custom-unsupported.xlsx");
        var failResult = await ProviderDrivers.Get(provider).CreateImporter()
            .ImportAsync(failSource, failRequest);

        Assert.False(reportResult.IsSuccess);
        Assert.Single(reportResult.Workbook.Rows);
        Assert.Contains(reportResult.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
        Assert.False(failResult.IsSuccess);
        Assert.Empty(failResult.Workbook.Rows);
        Assert.Contains(failResult.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
    }

    /// <summary>
    /// Decimal 规则应接受 Golden 输入中的边界内值。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task DecimalBetweenRule_ShouldAcceptGoldenValue(string provider)
    {
        using var source = Golden("validation/decimal-between-valid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
        Assert.Equal("5.5", result.Workbook.Rows[0].Code);
    }

    /// <summary>
    /// Date 规则应接受 Golden 输入中的边界内值。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task DateBetweenRule_ShouldAcceptGoldenValue(string provider)
    {
        using var source = Golden("validation/date-between-valid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
        Assert.Equal("2026-09-04", result.Workbook.Rows[0].Code);
    }

    /// <summary>
    /// Time 规则应接受 Golden 输入中的边界内值。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task TimeBetweenRule_ShouldAcceptGoldenValue(string provider)
    {
        using var source = Golden("validation/time-between-valid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Workbook.Rows);
        Assert.Equal("12:00:00", result.Workbook.Rows[0].Code);
    }

    /// <summary>
    /// 文本长度规则应拒绝 Golden 输入中的超长值。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public async Task TextLengthRule_ShouldRejectGoldenValue(string provider)
    {
        using var source = Golden("validation/text-length-invalid.xlsx");
        var result = await ProviderDrivers.Get(provider).CreateImporter().ImportAsync(source, CreateRequest());

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(nameof(ValidationContractRow.Code), error.PropertyName);
    }

    /// <summary>
    /// 验证 ClosedXML 同步和异步入口在原生规则先于配置规则的顺序下产生相同结果。
    /// </summary>
    [Fact]
    public async Task ClosedXmlConfiguredAndWorkbook_ShouldRunWorkbookRulesBeforeConfiguredRulesInBothModes()
    {
        using var source = CreateClosedXmlValidationWorkbook(string.Empty, XLAllowedValues.WholeNumber,
            value: "7", ignoreBlanks: false);
        var request = CreateRequest(validationMode: ExcelImportValidationMode.ConfiguredAndWorkbook);
        var importer = new ClosedXmlExcelImporter();

        var sync = importer.Import(source, request);
        source.Position = 0;
        var asynchronous = await importer.ImportAsync(source, request);

        Assert.False(sync.IsSuccess);
        Assert.False(asynchronous.IsSuccess);
        Assert.Empty(sync.Workbook.Rows);
        Assert.Empty(asynchronous.Workbook.Rows);
        var syncError = Assert.Single(sync.Errors);
        var asyncError = Assert.Single(asynchronous.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, syncError.Code);
        Assert.Equal(syncError.Code, asyncError.Code);
        Assert.Equal(syncError.Message, asyncError.Message);
        Assert.Equal(syncError.RowIndex, asyncError.RowIndex);
        Assert.Equal(syncError.PropertyName, asyncError.PropertyName);
    }

    /// <summary>
    /// 验证不可解析的名称列表按 UnsupportedFeaturePolicy 的 Report/Fail 契约处理。
    /// </summary>
    [Fact]
    public async Task ClosedXmlUnresolvableListReference_ShouldRespectReportAndFailPolicy()
    {
        var reportRequest = CreateRequest(ExcelUnsupportedFeaturePolicy.Report);
        var failRequest = CreateRequest(ExcelUnsupportedFeaturePolicy.Fail);
        var importer = new ClosedXmlExcelImporter();

        using var reportSource = CreateClosedXmlValidationWorkbook("Candidate", XLAllowedValues.List,
            value: "KnownValues", ignoreBlanks: false);
        var report = await importer.ImportAsync(reportSource, reportRequest);
        using var failSource = CreateClosedXmlValidationWorkbook("Candidate", XLAllowedValues.List,
            value: "KnownValues", ignoreBlanks: false);
        var fail = await importer.ImportAsync(failSource, failRequest);

        Assert.False(report.IsSuccess);
        Assert.Single(report.Workbook.Rows);
        Assert.Contains(report.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
        Assert.False(fail.IsSuccess);
        Assert.Empty(fail.Workbook.Rows);
        Assert.Contains(fail.Errors, error => error.Code == ExcelImportErrorCode.WorkbookValidation);
    }

    /// <summary>
    /// 创建 WorkbookRules 请求，禁用配置级 Required 校验。
    /// </summary>
    /// <param name="unsupportedFeaturePolicy">遇到不支持的工作簿规则时的处理策略。</param>
    /// <param name="validationMode">导入使用的规则来源。</param>
    /// <returns>使用继续收集错误模式读取 Data 工作表的请求。</returns>
    private static ExcelWorkbookImportRequest<ValidationContractWorkbook> CreateRequest(
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy = ExcelUnsupportedFeaturePolicy.Fail,
        ExcelImportValidationMode validationMode = ExcelImportValidationMode.WorkbookRules) =>
        ExcelImport.Workbook<ValidationContractWorkbook>(builder => builder
            .ValidationMode(validationMode)
            .UnsupportedFeaturePolicy(unsupportedFeaturePolicy)
            .Sheet<ValidationContractRow>("Data", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

    /// <summary>
    /// 创建包含真实 ClosedXML 原生数据校验的最小 XLSX 输入。
    /// </summary>
    /// <param name="code">待验证的 Code 单元格值。</param>
    /// <param name="allowedValues">工作簿原生数据校验类型。</param>
    /// <param name="value">等值比较使用的规则边界文本。</param>
    /// <param name="ignoreBlanks">是否忽略空单元格。</param>
    /// <returns>位于起点的工作簿内存流，由调用方释放。</returns>
    private static MemoryStream CreateClosedXmlValidationWorkbook(string code, XLAllowedValues allowedValues,
        string value, bool ignoreBlanks)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ValidationContractRow.Code);
            sheet.Cell("B1").Value = nameof(ValidationContractRow.Quantity);
            sheet.Cell("A2").Value = code;
            sheet.Cell("B2").Value = 1;
            var validation = sheet.Range("A2").CreateDataValidation();
            validation.AllowedValues = allowedValues;
            validation.Operator = XLOperator.EqualTo;
            validation.Value = value;
            validation.IgnoreBlanks = ignoreBlanks;
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 读取复制到测试输出目录的冻结 Golden 输入。
    /// </summary>
    /// <param name="relativePath">相对于 Golden 资源目录的样本路径。</param>
    /// <returns>包含完整样本内容的只读内存流，由调用方释放。</returns>
    private static MemoryStream Golden(string relativePath)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Resources", "Golden",
            relativePath.Replace('/', Path.DirectorySeparatorChar));
        Assert.True(File.Exists(path), $"Golden input was not copied: {path}");
        return new MemoryStream(File.ReadAllBytes(path), writable: false);
    }
}
