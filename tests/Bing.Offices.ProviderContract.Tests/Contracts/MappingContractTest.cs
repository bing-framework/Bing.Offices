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
/// 验证三套 Provider 的请求级映射和值映射公共合同。
/// </summary>
public sealed class MappingContractTest
{
    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 请求级映射结果应符合独立快照。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task MappingContract_ShouldMatchIndependentSnapshot(string provider)
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Mapping",
            ContractRequests.MappingExport(),
            ContractRequests.MappingImport());

        ProviderContractAssertions.HasNoException(execution);
        ProviderContractAssertions.HasResult(execution);
        Assert.True(execution.Result.IsSuccess,
            $"{execution.TraceId}: {string.Join(";", execution.Result.Errors.Select(error => error.Message))}");

        var actual = ContractSnapshots.MappingRows(execution.Result.Workbook);
        var expected = ContractSnapshots.ExpectedMappingRows();
        Assert.True(ContractSnapshotComparer.Equal(expected, actual),
            $"{execution.TraceId}: {ContractSnapshotComparer.Difference(expected, actual)}");
    }
}
