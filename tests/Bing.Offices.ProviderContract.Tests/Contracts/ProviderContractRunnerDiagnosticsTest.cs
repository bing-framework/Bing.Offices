using System.Threading.Tasks;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Runner;
using Bing.Offices.Testing.Requests;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证共享执行器提供完整的失败诊断上下文。
/// </summary>
public sealed class ProviderContractRunnerDiagnosticsTest
{
    /// <summary>
    /// 成功执行也必须保留场景、格式、预期、能力和实际结果。
    /// </summary>
    [Fact]
    public async Task SuccessfulExecution_ShouldExposeStableDiagnosticContext()
    {
        var driver = ProviderDrivers.Get("NPOI");
        var execution = await ProviderContractRunner.RoundTripAsync(
            driver,
            "Scalar",
            ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport());

        Assert.Equal("Scalar", execution.Scenario);
        Assert.Equal("NPOI", execution.Provider);
        Assert.Equal("Xlsx", execution.Format);
        Assert.Equal("Supported", execution.ExpectedProfile);
        Assert.Equal(driver.DeclaredCapabilities.Capabilities, execution.DeclaredCapabilities);
        Assert.Equal(ContractActualOutcome.Success, execution.ActualOutcome);
        Assert.Null(execution.Exception);
        Assert.Contains("scenario=Scalar", execution.DiagnosticContext);
        Assert.Contains("provider=NPOI", execution.DiagnosticContext);
        Assert.Contains("mode=Async", execution.DiagnosticContext);
        Assert.Contains("format=Xlsx", execution.DiagnosticContext);
        Assert.Contains("expectedProfile=Supported", execution.DiagnosticContext);
        Assert.Contains("declaredCapabilities=", execution.DiagnosticContext);
        Assert.Contains("actualOutcome=Success", execution.DiagnosticContext);
    }

    /// <summary>
    /// 故意不匹配时，统一断言消息必须保留全部七项诊断字段。
    /// </summary>
    [Fact]
    public async Task IntentionalMismatch_ShouldIncludeAllDiagnosticFields()
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get("NPOI"),
            "Scalar",
            ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport());

        var exception = Assert.ThrowsAny<System.Exception>(() =>
            Assert.True(false, execution.FormatAssertionFailure("intentional mismatch")));

        Assert.Contains("intentional mismatch", exception.Message);
        Assert.Contains("scenario=Scalar", exception.Message);
        Assert.Contains("provider=NPOI", exception.Message);
        Assert.Contains("mode=Async", exception.Message);
        Assert.Contains("format=Xlsx", exception.Message);
        Assert.Contains("expectedProfile=Supported", exception.Message);
        Assert.Contains("declaredCapabilities=", exception.Message);
        Assert.Contains("actualOutcome=Success", exception.Message);
    }

    /// <summary>
    /// 共享断言入口失败时也必须通过统一格式化器提供全部诊断字段。
    /// </summary>
    [Fact]
    public async Task SharedAssertionFailure_ShouldIncludeAllDiagnosticFields()
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get("NPOI"),
            "Scalar",
            ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport());
        var failed = new ProviderContractExecution<Bing.Offices.Testing.Models.ScalarContractWorkbook>(
            execution.TraceId, execution.Scenario, execution.Provider, execution.Mode, execution.Format,
            execution.ExpectedProfile, execution.DeclaredCapabilities, ContractActualOutcome.Exception,
            execution.Output, execution.Result, new System.InvalidOperationException("intentional helper failure"));

        var exception = Assert.ThrowsAny<System.Exception>(() => ProviderContractAssertions.HasNoException(failed));

        Assert.Contains("unexpected exception=System.InvalidOperationException", exception.Message);
        Assert.Contains("scenario=Scalar", exception.Message);
        Assert.Contains("provider=NPOI", exception.Message);
        Assert.Contains("mode=Async", exception.Message);
        Assert.Contains("format=Xlsx", exception.Message);
        Assert.Contains("expectedProfile=Supported", exception.Message);
        Assert.Contains("declaredCapabilities=", exception.Message);
        Assert.Contains("actualOutcome=Exception", exception.Message);
    }
}
