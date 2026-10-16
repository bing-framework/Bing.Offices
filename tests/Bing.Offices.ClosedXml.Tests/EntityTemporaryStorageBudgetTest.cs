using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Tests;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML 前向实体临时存储预算合同。
/// </summary>
public sealed class EntityTemporaryStorageBudgetTest : EntityTemporaryStorageBudgetContractTest
{
    protected override string ProviderName => "ClosedXML";
    protected override ServiceProvider CreateProvider() =>
        new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
}
