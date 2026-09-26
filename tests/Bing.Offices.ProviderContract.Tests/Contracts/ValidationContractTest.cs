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
/// 验证结构化校验错误的共同可观察结果。
/// </summary>
public sealed class ValidationContractTest
{
    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });

    /// <summary>
    /// 必填校验失败应返回结构化错误且不提交无效实体。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task ValidationContract_ShouldReturnStableErrorSnapshot(string provider)
    {
        var execution = await ProviderContractRunner.RoundTripAsync(
            ProviderDrivers.Get(provider),
            "Validation",
            ContractRequests.ValidationExport(),
            ContractRequests.ValidationImport());

        ProviderContractAssertions.HasNoException(execution);
        ProviderContractAssertions.HasResult(execution);
        Assert.False(execution.Result.IsSuccess, execution.FormatAssertionFailure("validation unexpectedly succeeded"));
        Assert.True(!execution.Result.Workbook.Rows.Any(),
            execution.FormatAssertionFailure("validation retained rejected entities"));
        var errors = ContractSnapshots.Errors(execution.Result.Errors);
        Assert.True(errors.Any(), execution.FormatAssertionFailure("validation did not return errors"));
        Assert.True(errors.All(error => error.SheetName == "Data"),
            execution.FormatAssertionFailure("validation error sheet differs from the contract"));
        Assert.True(errors.All(error => error.RowIndex == 2),
            execution.FormatAssertionFailure("validation error row differs from the contract"));
        Assert.True(errors.Any(error => error.PropertyName == "Code"),
            execution.FormatAssertionFailure("validation error does not identify Code"));
    }
}
