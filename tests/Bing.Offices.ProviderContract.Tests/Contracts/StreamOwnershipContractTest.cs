using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Bing.Offices.Testing.Requests;
using Bing.Offices.Testing.Streams;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 Provider 不拥有调用方提供的源和目标流。
/// </summary>
public sealed class StreamOwnershipContractTest
{
    /// <summary>
    /// 异步导出和导入结束后调用方流应保持可用。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    [Theory]
    [MemberData(nameof(Providers))]
    public async Task AsyncRoundTrip_ShouldLeaveCallerStreamsOpen(string provider)
    {
        var driver = ProviderDrivers.Get(provider);
        using var innerOutput = new MemoryStream();
        using var output = new TrackingStream(innerOutput, leaveInnerOpen: true);
        await driver.CreateExporter().ExportAsync(ContractRequests.ScalarExport(), output);
        Assert.False(output.WasDisposed);
        Assert.True(output.BytesWritten > 0);
        var bytes = innerOutput.ToArray();

        using var innerInput = new MemoryStream(bytes, writable: false);
        using var input = new TrackingStream(innerInput, leaveInnerOpen: true);
        var result = await driver.CreateImporter().ImportAsync(input, ContractRequests.ScalarImport());

        Assert.False(input.WasDisposed);
        Assert.True(input.BytesRead > 0);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
    }

    /// <summary>
    /// 获取全部 Provider 名称。
    /// </summary>
    public static IEnumerable<object[]> Providers => ProviderDrivers.All.Select(driver => new object[] { driver.Name });
}
