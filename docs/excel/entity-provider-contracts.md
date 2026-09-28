# Entity Layout Provider 合同测试包

`Bing.Offices.EntityProviderContracts` 提供可复用的 Entity Layout 合同测试程序集。它只依赖公开的 Abstractions，不依赖 xUnit 或官方 Provider。第三方测试工程引用该包后，将 Provider 的公开 Entity 导入、导出接口传给合同入口。仓库只在本地与 CI 生成包，尚未发布到公共 NuGet 源。

从基础 XLSX 工作簿导出开始开发 Provider，可参考[独立项目模板](../../templates/excel-provider-starter/README.md)。模板尚未实现 Entity Layout，完成对应接口后再启用本文合同。

包名为 `Bing.Offices.EntityProviderContracts`，类型的命名空间仍为 `Bing.Offices.Testing.Entities`，以便已有源码链接调用保持兼容。

## 接入方式

先运行 `dotnet pack tests/Bing.Offices.EntityProviderContracts/Bing.Offices.EntityProviderContracts.csproj -c Release -o artifacts/packages-contracts`，把产物目录加入测试工程的 NuGet 源。包版本沿用当前构建版本；本地示例为：

```xml
<ItemGroup>
  <PackageReference Include="Bing.Offices.EntityProviderContracts" Version="2.0.0" PrivateAssets="all" />
</ItemGroup>
```

测试工程还需要引用待测 Provider。合同包依赖包含 `IExcelEntityExporter`、`IExcelEntityImporter` 与 Entity Layout API 的 Abstractions。调用时创建 Provider 的公开服务并传入合同入口：

```csharp
using Bing.Offices.Entities;
using Bing.Offices.Testing.Entities;
using Xunit;

public sealed class EntityLayoutProviderContractTest
{
    [Fact]
    public void CoreLayout_ShouldRoundTrip()
    {
        IExcelEntityExporter exporter = CreateProviderExporter();
        IExcelEntityImporter importer = CreateProviderImporter();

        EntityLayoutProviderContractSuite.VerifyCore(exporter, importer);
    }

    [Fact]
    public void PageSubtotal_ShouldRoundTrip()
    {
        EntityLayoutProviderContractSuite.VerifyPageSubtotal(
            CreateProviderExporter(), CreateProviderImporter());
    }

    [Fact]
    public void GroupSubtotal_ShouldRoundTrip()
    {
        EntityLayoutProviderContractSuite.VerifyGroupSubtotal(
            CreateProviderExporter(), CreateProviderImporter());
    }

    [Fact]
    public void FooterFormulas_ShouldRoundTrip()
    {
        EntityLayoutProviderContractSuite.VerifyFooterFormulas(
            CreateProviderExporter(), CreateProviderImporter());
    }

    [Fact]
    public void AttributeAndNamedAnchors_ShouldRoundTripTemplate()
    {
        EntityLayoutProviderContractSuite.VerifyAttributeAndNamedAnchors(
            CreateProviderExporter(), CreateProviderImporter());
    }

    [Fact]
    public void NamedListFixedCellCollision_ShouldFailBeforeOutput()
    {
        EntityLayoutProviderContractSuite.VerifyNamedListFixedCellCollision(
            CreateProviderExporter(), CreateProviderImporter());
    }

    [Fact]
    public async Task AsyncLayout_ShouldRoundTrip()
    {
        await EntityLayoutProviderContractSuite.VerifyAsync(
            CreateProviderExporter(), CreateProviderImporter());
    }

    private static IExcelEntityExporter CreateProviderExporter() =>
        throw new System.NotImplementedException();

    private static IExcelEntityImporter CreateProviderImporter() =>
        throw new System.NotImplementedException();
}
```

替换 `CreateProviderExporter` 和 `CreateProviderImporter` 的示意实现，并将每个入口放入测试框架的独立测试中。模板断言失败会抛出带场景名称的 `InvalidOperationException`；Provider 自身抛出的异常仍按原类型透传。

已有源码链接方式仍可使用：将 `tests/Bing.Offices.Testing/Entities/EntityLayoutProviderContractSuite.cs` 作为 `Compile` 项加入测试工程。不要同时链接源码和引用合同包，否则会定义两个同名类型。

## 合同范围

- `VerifyCore` 检查两个动态字典、最终 Footer 聚合、命名锚点和完整明细读回。
- `VerifyAttributeAndNamedAnchors` 检查属性式固定单元格、`CellNamed` 与 `ListRegionNamed` 在 XLSX 模板中的导入导出往返。
- `VerifyNamedListFixedCellCollision` 检查命名列表按模板实际地址与固定单元格冲突时的 Plan 配置错误及目标流保护。
- `VerifyPageSubtotal` 检查分页小计及最终合计的单元格值，并验证导入跳过小计行。
- `VerifyGroupSubtotal` 检查每个连续分组的小计及最终合计，并验证导入跳过小计行。
- `VerifyFooterFormulas` 检查显式 Footer 公式、连续明细求和及跨小计明细求和公式的单元格类型/公式文本、空与非空明细及导入边界；连续求和与中间小计冲突时在布局阶段拒绝。
- `VerifyAsync` 检查异步往返、预取消、目标内容保护和调用方流所有权。

七项入口以 XLSX 格式为前提。模板锚点合同需要 Provider 支持 Entity Layout 模板导入导出；分页小计、分组小计和 Footer 公式入口只对明确实现相应行为的 Provider 运行。不支持时，应单独测试公开能力声明和拒绝行为，不能通过跳过合同掩盖能力缺口。`VerifyFooterFormulas` 要求 Provider 将公式写入真实公式单元格，并支持 `FormulaSumContiguousRowsAbove` 与 `FormulaSumDetailRowsAbove` 的空/非空明细语义；它不验证公式引擎的计算结果。连续求和只适用于尾部上方连续的明细行，跨分页或分组小计时使用仅汇总实际明细行的 `FormulaSumDetailRowsAbove`。

此源码模板验证可移植的 Entity Layout 行为，不覆盖所有 Provider 特有行为。特别是 XLS/HSSF 的格式支持、边界和失败情形仍需由 Provider 自己的测试覆盖；NPOI 的 XLS 测试不能由这里的 XLSX 合同替代。
