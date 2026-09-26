using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.ProviderContract.Tests.Runner;
using Bing.Offices.Testing.Requests;
using Bing.Offices.Testing.Snapshots;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证同步与异步外层 IO 返回同一公共快照。
/// </summary>
public sealed class AsyncContractTest
{
    /// <summary>
    /// 同步和异步模式均应成功并产生相同的标量结果。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task SyncAndAsync_ShouldMatchScalarSnapshot(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        var sync = await ProviderContractRunner.RoundTripAsync(driver, "Async", ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport(), ContractExecutionMode.Sync);
        var asynchronous = await ProviderContractRunner.RoundTripAsync(driver, "Async", ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport(), ContractExecutionMode.Async);

        ProviderContractAssertions.HasNoException(sync);
        ProviderContractAssertions.HasNoException(asynchronous);
        Assert.True(ContractExecutionMode.Sync == sync.Mode,
            sync.FormatAssertionFailure("sync mode mismatch"));
        Assert.True(ContractExecutionMode.Async == asynchronous.Mode,
            asynchronous.FormatAssertionFailure("async mode mismatch"));
        Assert.True(sync.Result.IsSuccess, sync.FormatAssertionFailure(
            string.Join(";", sync.Result.Errors.Select(error => error.Message))));
        Assert.True(asynchronous.Result.IsSuccess,
            asynchronous.FormatAssertionFailure(
                string.Join(";", asynchronous.Result.Errors.Select(error => error.Message))));
        Assert.True(ContractSnapshotComparer.Equal(ContractSnapshots.ScalarRows(sync.Result.Workbook),
            ContractSnapshots.ScalarRows(asynchronous.Result.Workbook)),
            sync.FormatAssertionFailure("sync/async snapshot mismatch"));
        Assert.True(sync.Output.Length > 0, sync.FormatAssertionFailure("sync output is empty"));
        Assert.True(asynchronous.Output.Length > 0, asynchronous.FormatAssertionFailure("async output is empty"));
    }

    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
