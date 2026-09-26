using System;
using System.Linq;
using Bing.Offices.ProviderContract.Tests.Drivers;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Profiles;

/// <summary>
/// 保护 Provider 枚举、预期档案和执行入口的一致性。
/// </summary>
public sealed class ProfileCompletenessTest
{
    /// <summary>
    /// 每个驱动器必须有且只有一个独立预期档案。
    /// </summary>
    [Fact]
    public void EveryDriver_ShouldHaveIndependentCompleteProfile()
    {
        var drivers = ProviderDrivers.All;
        var profiles = ProviderContractProfiles.All;
        Assert.Equal(drivers.Count, profiles.Count);
        Assert.Equal(drivers.Select(driver => driver.Name).OrderBy(name => name),
            profiles.Select(profile => profile.Provider).OrderBy(name => name));

        foreach (var profile in profiles)
        {
            foreach (var scenario in Enum.GetValues(typeof(ContractScenario)).Cast<ContractScenario>())
                Assert.True(profile.Expectations.ContainsKey(scenario),
                    $"Missing {scenario} expectation for {profile.Provider}.");
        }
    }

    /// <summary>
    /// 档案预期不应由生产声明自动决定。
    /// </summary>
    [Fact]
    public void ExpectedProfile_ShouldBeSeparateFromDeclaredCapabilities()
    {
        var mini = ProviderContractProfiles.Get("MiniExcel");
        var driver = ProviderDrivers.Get("MiniExcel");
        Assert.Equal(ContractExpectation.NotApplicable, mini.For(ContractScenario.Entity));
        Assert.False(driver.DeclaredCapabilities.Supports(Bing.Offices.Providers.ExcelProviderCapabilities.Entity));
    }
}
