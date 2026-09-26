using System;
using System.Collections.Generic;

namespace Bing.Offices.ProviderContract.Tests.Profiles;

/// <summary>
/// 合同场景的稳定标识。
/// </summary>
public enum ContractScenario
{
    /// <summary>
    /// 标量和日期场景。
    /// </summary>
    Scalar,
    /// <summary>
    /// 动态列场景。
    /// </summary>
    DynamicColumns,
    /// <summary>
    /// 映射和值映射场景。
    /// </summary>
    Mapping,
    /// <summary>
    /// 结构化校验错误场景。
    /// </summary>
    Validation,
    /// <summary>
    /// 父子关系场景。
    /// </summary>
    Relations,
    /// <summary>
    /// 异步外围 IO 场景。
    /// </summary>
    Async,
    /// <summary>
    /// 公共单元格样式场景。
    /// </summary>
    Style,
    /// <summary>
    /// 公共合并区域场景。
    /// </summary>
    Merge,
    /// <summary>
    /// 公共模板保真场景。
    /// </summary>
    Template,
    /// <summary>
    /// 实体列表区域场景。
    /// </summary>
    Entity,
    /// <summary>
    /// 实体列表区域动态列场景。
    /// </summary>
    EntityDynamicColumns,
    /// <summary>
    /// 公式文本和缓存值场景。
    /// </summary>
    Formula
}

/// <summary>
/// 合同预期结果类别。
/// </summary>
public enum ContractExpectation
{
    /// <summary>
    /// 应成功并符合公共快照。
    /// </summary>
    Supported,
    /// <summary>
    /// 应以结构化不支持结果失败。
    /// </summary>
    UnsupportedExpected,
    /// <summary>
    /// 仅部分公共语义可用。
    /// </summary>
    Partial,
    /// <summary>
    /// 当前 Provider 不适用。
    /// </summary>
    NotApplicable
}

/// <summary>
/// 测试所有者维护的 Provider 预期档案。
/// </summary>
/// <remarks>预期结果独立维护，不从生产能力声明生成。</remarks>
public sealed class ProviderContractProfile
{
    /// <summary>
    /// 初始化一个 <see cref="ProviderContractProfile"/> 类型的实例。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="expectations">场景与预期结果的映射。</param>
    public ProviderContractProfile(string provider, IReadOnlyDictionary<ContractScenario, ContractExpectation> expectations)
    {
        Provider = provider ?? throw new ArgumentNullException(nameof(provider));
        Expectations = expectations ?? throw new ArgumentNullException(nameof(expectations));
    }

    /// <summary>
    /// 获取 Provider 名称。
    /// </summary>
    public string Provider { get; }
    /// <summary>
    /// 获取场景预期映射。
    /// </summary>
    public IReadOnlyDictionary<ContractScenario, ContractExpectation> Expectations { get; }

    /// <summary>
    /// 获取指定场景预期。
    /// </summary>
    /// <param name="scenario">合同场景。</param>
    /// <returns>该场景的预期结果。</returns>
    public ContractExpectation For(ContractScenario scenario) =>
        Expectations.TryGetValue(scenario, out var value)
            ? value
            : throw new InvalidOperationException($"Profile {Provider} has no expectation for {scenario}.");
}

/// <summary>
/// 独立维护的 Provider 预期档案集合。
/// </summary>
public static class ProviderContractProfiles
{
    /// <summary>
    /// 获取所有 Provider 预期档案。
    /// </summary>
    public static IReadOnlyList<ProviderContractProfile> All { get; } = new[]
    {
        Create("NPOI", entityDynamic: ContractExpectation.UnsupportedExpected),
        Create("MiniExcel", rich: ContractExpectation.UnsupportedExpected,
            entity: ContractExpectation.NotApplicable,
            entityDynamic: ContractExpectation.NotApplicable),
        Create("ClosedXML")
    };

    /// <summary>
    /// 按名称获取预期档案。
    /// </summary>
    /// <param name="provider">区分大小写的 Provider 名称。</param>
    /// <returns>名称匹配的预期档案。</returns>
    public static ProviderContractProfile Get(string provider)
    {
        foreach (var profile in All)
            if (string.Equals(profile.Provider, provider, StringComparison.Ordinal))
                return profile;
        throw new ArgumentException($"Unknown provider contract profile: {provider}", nameof(provider));
    }

    /// <summary>
    /// 创建 Provider 各合同场景的预期档案。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="rich">样式、合并和模板场景的预期结果。</param>
    /// <param name="entity">实体场景的预期结果。</param>
    /// <param name="entityDynamic">实体动态列场景的预期结果。</param>
    /// <returns>包含全部合同场景预期的档案。</returns>
    private static ProviderContractProfile Create(string provider,
        ContractExpectation rich = ContractExpectation.Supported,
        ContractExpectation entity = ContractExpectation.Supported,
        ContractExpectation entityDynamic = ContractExpectation.Supported)
    {
        return new ProviderContractProfile(provider, new Dictionary<ContractScenario, ContractExpectation>
        {
            [ContractScenario.Scalar] = ContractExpectation.Supported,
            [ContractScenario.DynamicColumns] = ContractExpectation.Supported,
            [ContractScenario.Mapping] = ContractExpectation.Supported,
            [ContractScenario.Validation] = ContractExpectation.Supported,
            [ContractScenario.Relations] = ContractExpectation.Supported,
            [ContractScenario.Async] = ContractExpectation.Supported,
            [ContractScenario.Style] = rich,
            [ContractScenario.Merge] = rich,
            [ContractScenario.Template] = rich,
            [ContractScenario.Entity] = entity,
            [ContractScenario.EntityDynamicColumns] = entityDynamic,
            [ContractScenario.Formula] = string.Equals(provider, "ClosedXML", StringComparison.Ordinal)
                ? ContractExpectation.Supported
                : ContractExpectation.Partial
        });
    }
}
