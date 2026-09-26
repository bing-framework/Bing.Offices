using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Profiles;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Runner;

/// <summary>
/// 合同执行模式。
/// </summary>
public enum ContractExecutionMode
{
    /// <summary>
    /// 同步导出和导入。
    /// </summary>
    Sync,
    /// <summary>
    /// 异步导出和导入。
    /// </summary>
    Async
}

/// <summary>
/// 合同执行的稳定可观察结果类别。
/// </summary>
public enum ContractActualOutcome
{
    /// <summary>
    /// 导出和导入均成功。
    /// </summary>
    Success,
    /// <summary>
    /// 导入完成但返回结构化失败结果。
    /// </summary>
    ResultFailure,
    /// <summary>
    /// 执行抛出异常。
    /// </summary>
    Exception,
    /// <summary>
    /// 执行未产生导入结果。
    /// </summary>
    NoResult
}

/// <summary>
/// 单次 Provider 合同运行的结果与诊断上下文。
/// </summary>
/// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
public sealed class ProviderContractExecution<TWorkbook> where TWorkbook : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ProviderContractExecution{TWorkbook}" /> 类型的实例。
    /// </summary>
    /// <param name="traceId">关联本次执行日志的追踪标识。</param>
    /// <param name="scenario">合同场景名称。</param>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">同步或异步执行模式。</param>
    /// <param name="format">导出文件格式名称。</param>
    /// <param name="expectedProfile">独立预期档案中的结果类别。</param>
    /// <param name="declaredCapabilities">Provider 声明的能力组合。</param>
    /// <param name="actualOutcome">本次实际结果类别。</param>
    /// <param name="output">执行期间写出的字节。</param>
    /// <param name="result">导入结果；未得到结果时为空。</param>
    /// <param name="exception">执行捕获的异常；未发生异常时为空。</param>
    internal ProviderContractExecution(string traceId, string scenario, string provider,
        ContractExecutionMode mode, string format, string expectedProfile,
        ExcelProviderCapabilities declaredCapabilities, ContractActualOutcome actualOutcome,
        byte[] output, ExcelWorkbookImportResult<TWorkbook> result, Exception exception)
    {
        TraceId = traceId;
        Scenario = scenario;
        Provider = provider;
        Mode = mode;
        Format = format;
        ExpectedProfile = expectedProfile;
        DeclaredCapabilities = declaredCapabilities;
        ActualOutcome = actualOutcome;
        Output = output;
        Result = result;
        Exception = exception;
    }

    /// <summary>
    /// 获取可关联日志的追踪标识。
    /// </summary>
    public string TraceId { get; }
    /// <summary>
    /// 获取合同场景标识。
    /// </summary>
    public string Scenario { get; }
    /// <summary>
    /// 获取 Provider 名称。
    /// </summary>
    public string Provider { get; }
    /// <summary>
    /// 获取执行模式。
    /// </summary>
    public ContractExecutionMode Mode { get; }
    /// <summary>
    /// 获取导出请求使用的文件格式。
    /// </summary>
    public string Format { get; }
    /// <summary>
    /// 获取独立维护的预期档案结果。
    /// </summary>
    public string ExpectedProfile { get; }
    /// <summary>
    /// 获取 Provider 声明的能力集合。
    /// </summary>
    public ExcelProviderCapabilities DeclaredCapabilities { get; }
    /// <summary>
    /// 获取本次执行的稳定实际结果类别。
    /// </summary>
    public ContractActualOutcome ActualOutcome { get; }
    /// <summary>
    /// 获取导出输出字节。
    /// </summary>
    public byte[] Output { get; }
    /// <summary>
    /// 获取导入结果。
    /// </summary>
    /// <remarks>结果可包含结构化失败信息；未得到结果时为空。</remarks>
    public ExcelWorkbookImportResult<TWorkbook> Result { get; }
    /// <summary>
    /// 获取失败时捕获的异常。
    /// </summary>
    public Exception Exception { get; }
    /// <summary>
    /// 获取是否得到正常导入结果。
    /// </summary>
    public bool HasResult => Result != null;

    /// <summary>
    /// 获取可直接附加到合同断言失败消息的完整诊断上下文。
    /// </summary>
    public string DiagnosticContext => $"scenario={Scenario}; provider={Provider}; mode={Mode}; format={Format}; "
        + $"expectedProfile={ExpectedProfile}; declaredCapabilities={DeclaredCapabilities}; actualOutcome={ActualOutcome}";

    /// <summary>
    /// 将合同断言说明补全为稳定的七字段诊断上下文。
    /// </summary>
    /// <param name="message">断言的业务说明。</param>
    /// <returns>可直接传递给断言框架的完整失败消息。</returns>
    public string FormatAssertionFailure(string message) => $"{message}; {DiagnosticContext}";
}

/// <summary>
/// 统一执行 Provider 合同往返流程。
/// </summary>
public static class ProviderContractRunner
{
    /// <summary>
    /// 执行 Provider 导出导入往返并捕获结果。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="driver">创建导入导出器并提供能力声明的驱动器。</param>
    /// <param name="scenario">合同场景名称。</param>
    /// <param name="exportRequest">工作簿导出请求。</param>
    /// <param name="importRequest">工作簿导入请求。</param>
    /// <param name="mode">同步或异步执行模式。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果包含导入结果、输出字节和捕获的异常。</returns>
    public static async Task<ProviderContractExecution<TWorkbook>> RoundTripAsync<TWorkbook>(
        IProviderContractDriver driver,
        string scenario,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<TWorkbook> importRequest,
        ContractExecutionMode mode = ContractExecutionMode.Async,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (driver == null) throw new ArgumentNullException(nameof(driver));
        return await RoundTripAsync(driver.Name, scenario, driver.CreateExporter(), driver.CreateImporter(),
            driver.DeclaredCapabilities.Capabilities, exportRequest, importRequest, mode, cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// 执行 Provider 导出导入往返并捕获结果。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="scenario">合同场景名称。</param>
    /// <param name="exporter">待验证的导出器。</param>
    /// <param name="importer">待验证的导入器。</param>
    /// <param name="exportRequest">工作簿导出请求。</param>
    /// <param name="importRequest">工作簿导入请求。</param>
    /// <param name="mode">同步或异步执行模式。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果包含导入结果、输出字节和捕获的异常。</returns>
    public static Task<ProviderContractExecution<TWorkbook>> RoundTripAsync<TWorkbook>(
        string provider,
        string scenario,
        IExcelExporter exporter,
        IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<TWorkbook> importRequest,
        ContractExecutionMode mode = ContractExecutionMode.Async,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new() =>
        RoundTripAsync(provider, scenario, exporter, importer,
            GetDeclaredCapabilities(exporter, importer), exportRequest, importRequest, mode, cancellationToken);

    /// <summary>
    /// 执行 Provider 导出导入往返并捕获结果。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="scenario">合同场景名称。</param>
    /// <param name="exporter">待验证的导出器。</param>
    /// <param name="importer">待验证的导入器。</param>
    /// <param name="declaredCapabilities">用于诊断上下文的 Provider 能力声明。</param>
    /// <param name="exportRequest">工作簿导出请求。</param>
    /// <param name="importRequest">工作簿导入请求。</param>
    /// <param name="mode">同步或异步执行模式。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步操作，结果包含导入结果、输出字节和捕获的异常。</returns>
    public static async Task<ProviderContractExecution<TWorkbook>> RoundTripAsync<TWorkbook>(
        string provider,
        string scenario,
        IExcelExporter exporter,
        IExcelImporter importer,
        ExcelProviderCapabilities declaredCapabilities,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<TWorkbook> importRequest,
        ContractExecutionMode mode = ContractExecutionMode.Async,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (provider == null) throw new ArgumentNullException(nameof(provider));
        if (scenario == null) throw new ArgumentNullException(nameof(scenario));
        if (exporter == null) throw new ArgumentNullException(nameof(exporter));
        if (importer == null) throw new ArgumentNullException(nameof(importer));

        using var output = new MemoryStream();
        ExcelWorkbookImportResult<TWorkbook> result = null;
        Exception exception = null;
        try
        {
            if (mode == ContractExecutionMode.Sync)
            {
                exporter.Export(exportRequest, output, cancellationToken);
                output.Position = 0;
                result = importer.Import(output, importRequest, cancellationToken);
            }
            else
            {
                await exporter.ExportAsync(exportRequest, output, cancellationToken)
                    .ConfigureAwait(false);
                output.Position = 0;
                result = await importer.ImportAsync(output, importRequest, cancellationToken)
                    .ConfigureAwait(false);
            }
        }
        catch (Exception caught)
        {
            exception = caught;
        }

        var actualOutcome = exception != null
            ? ContractActualOutcome.Exception
            : result == null
                ? ContractActualOutcome.NoResult
                : result.IsSuccess
                    ? ContractActualOutcome.Success
                    : ContractActualOutcome.ResultFailure;
        return new ProviderContractExecution<TWorkbook>(
            $"scenario={scenario}; provider={provider}; mode={mode}; format={exportRequest?.Format ?? default}; "
            + $"expectedProfile={ResolveExpectedProfile(provider, scenario)}; declaredCapabilities={declaredCapabilities}; "
            + $"actualOutcome={actualOutcome}", scenario, provider, mode,
            exportRequest?.Format.ToString() ?? "Unknown",
            ResolveExpectedProfile(provider, scenario), declaredCapabilities, actualOutcome,
            output.ToArray(), result, exception);
    }

    /// <summary>
    /// 解析合同场景的预期结果名称。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="scenario">合同场景名称。</param>
    /// <returns>预期结果名称；无法识别场景时返回 Unspecified。</returns>
    private static string ResolveExpectedProfile(string provider, string scenario)
    {
        if (string.Equals(scenario, "Cancellation", StringComparison.Ordinal))
            return "OperationCanceled";
        if (string.Equals(scenario, "Exception", StringComparison.Ordinal))
            return "StructuredException";

        var normalizedScenario = scenario switch
        {
            "Converter" => nameof(ContractScenario.Mapping),
            "Relation" => nameof(ContractScenario.Relations),
            _ => scenario
        };
        if (!Enum.TryParse(normalizedScenario, ignoreCase: true, out ContractScenario parsed))
            return "Unspecified";
        return ProviderContractProfiles.Get(provider).For(parsed).ToString();
    }

    /// <summary>
    /// 读取导入导出器声明的能力组合。
    /// </summary>
    /// <param name="exporter">优先读取能力的导出器。</param>
    /// <param name="importer">导出器未声明能力时读取的导入器。</param>
    /// <returns>首个可用能力声明；两者均未声明时返回 None。</returns>
    private static ExcelProviderCapabilities GetDeclaredCapabilities(
        IExcelExporter exporter, IExcelImporter importer)
    {
        if (exporter is IExcelProviderCapabilities exporterCapabilities)
            return exporterCapabilities.Capabilities;
        if (importer is IExcelProviderCapabilities importerCapabilities)
            return importerCapabilities.Capabilities;
        return ExcelProviderCapabilities.None;
    }
}

/// <summary>
/// Provider 合同的共享断言入口。
/// </summary>
/// <remarks>失败消息包含运行器的完整诊断上下文。</remarks>
public static class ProviderContractAssertions
{
    /// <summary>
    /// 断言本次执行未捕获异常。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="execution">待断言的合同执行结果。</param>
    public static void HasNoException<TWorkbook>(ProviderContractExecution<TWorkbook> execution)
        where TWorkbook : class, new()
        => Assert.True(execution.Exception == null, execution.FormatAssertionFailure(
            $"unexpected exception={execution.Exception?.GetType().FullName}: {execution.Exception?.Message}"));

    /// <summary>
    /// 断言本次执行产生了导入结果。
    /// </summary>
    /// <typeparam name="TWorkbook">导入工作簿模型类型。</typeparam>
    /// <param name="execution">待断言的合同执行结果。</param>
    public static void HasResult<TWorkbook>(ProviderContractExecution<TWorkbook> execution)
        where TWorkbook : class, new()
        => Assert.True(execution.Result != null, execution.FormatAssertionFailure("import result is null"));
}
