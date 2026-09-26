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
/// 驱动三套 Provider 执行相同标量输入并匹配独立快照。
/// </summary>
public sealed class ScalarContractTest
{
    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 标量、日期、可空值和动态值应符合共同预期。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task ScalarContract_ShouldMatchIndependentSnapshot(string provider)
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Scalar",
            ContractRequests.ScalarExport(),
            ContractRequests.ScalarImport());

        ProviderContractAssertions.HasNoException(execution);
        ProviderContractAssertions.HasResult(execution);
        Assert.True(execution.Result.IsSuccess,
            $"{execution.TraceId}: {string.Join(";", execution.Result.Errors.Select(error => error.Message))}");

        var actual = ContractSnapshots.ScalarRows(execution.Result.Workbook);
        var expected = ContractSnapshots.ExpectedScalarRows();
        Assert.True(ContractSnapshotComparer.Equal(expected, actual),
            $"{execution.TraceId}: {ContractSnapshotComparer.Difference(expected, actual)}");
    }
}
