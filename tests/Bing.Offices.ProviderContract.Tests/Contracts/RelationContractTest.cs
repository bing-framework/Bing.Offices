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
/// 验证 Workbook 级父子关系公共合同。
/// </summary>
public sealed class RelationContractTest
{
    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 父子关系应按不区分大小写的键绑定并保持输入顺序。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task RelationContract_ShouldMatchIndependentSnapshot(string provider)
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Relations",
            ContractRequests.RelationExport(),
            ContractRequests.RelationImport());

        ProviderContractAssertions.HasNoException(execution);
        ProviderContractAssertions.HasResult(execution);
        Assert.True(execution.Result.IsSuccess,
            $"{execution.TraceId}: {string.Join(";", execution.Result.Errors.Select(error => error.Message))}");

        var actual = ContractSnapshots.Relations(execution.Result.Workbook);
        var expected = new[]
        {
            new RelationSnapshot("A-1", new[] { "Item-1", "Item-2" })
        };
        Assert.True(ContractSnapshotComparer.Equal(expected, actual),
            $"{execution.TraceId}: {ContractSnapshotComparer.Difference(expected, actual)}");
    }
}
