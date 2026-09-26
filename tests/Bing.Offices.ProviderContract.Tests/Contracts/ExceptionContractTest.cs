using System.Collections.Generic;
using System.IO;
using System.Linq;
using System;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Runner;
using Bing.Offices.Testing.Models;
using Bing.Offices.Testing.Requests;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证结构化资源异常可由共享执行器稳定观察。
/// </summary>
public sealed class ExceptionContractTest
{
    /// <summary>
    /// 输入字节限制超出时应保留结构化异常和失败上下文。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task ResourceLimit_ShouldExposeStructuredExceptionMetadata(string provider)
    {
        var request = ExcelImport.Workbook<ScalarContractWorkbook>(builder => builder
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
            .Sheet<ScalarContractRow>("Data", root => root.Rows));

        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Exception",
            ContractRequests.ScalarExport(),
            request);

        Assert.True(execution.Exception is BingOfficesResourceLimitException,
            execution.FormatAssertionFailure("resource limit exception type mismatch"));
        var exception = (BingOfficesResourceLimitException)execution.Exception;
        Assert.True(ContractActualOutcome.Exception == execution.ActualOutcome,
            execution.FormatAssertionFailure("exception outcome mismatch"));
        Assert.True("StructuredException" == execution.ExpectedProfile,
            execution.FormatAssertionFailure("exception profile mismatch"));
        Assert.True("Xlsx" == execution.Format, execution.FormatAssertionFailure("format mismatch"));
        Assert.True(BingOfficesErrorCode.ResourceLimitExceeded == exception.Code,
            execution.FormatAssertionFailure("exception code mismatch"));
        Assert.True(provider == exception.Provider, execution.FormatAssertionFailure("exception provider mismatch"));
        Assert.True(BingOfficesOperation.Import == exception.Operation,
            execution.FormatAssertionFailure("exception operation mismatch"));
        Assert.True(Enum.IsDefined(typeof(BingOfficesStage), exception.Stage),
            execution.FormatAssertionFailure("exception stage is undefined"));
        Assert.True(execution.Output.Length > 0, execution.FormatAssertionFailure("output is empty"));
        Assert.True(execution.Result == null, execution.FormatAssertionFailure("unexpected import result"));
        Assert.Contains("scenario=Exception", execution.DiagnosticContext);
        Assert.Contains($"provider={provider}", execution.DiagnosticContext);
    }

    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
