# Continuation 12：可复用 Entity Layout Provider 合同模板执行记录

## 结果

本阶段 COMPLETED。新增单文件 `EntityLayoutProviderContractSuite.cs`，仅依赖公开 Entity 导入/导出接口和标准 .NET API，不依赖 xUnit 或任何官方 Provider 程序集。四个独立入口核对动态字典与最终 Footer、分页小计、连续分组小计，以及异步往返、预取消和流所有权。聚合值通过公开固定单元格导入再读回，最终 Footer 名称通过 XLSX 元数据验证；ClosedXML 的 `A4:A4` 与 NPOI 的 `A4` 均按单格名称处理。

仓库 Provider Contract 对 NPOI 和 ClosedXML 实际调用同一模板；独立第三方包消费者通过源码链接该文件，只引用生成的 NuGet 包并在 net6.0、net8.0 执行。测试模板不是已发布 NuGet 包，未修改生产 API、Provider 能力声明或版本号。NPOI XLS/HSSF 专属矩阵继续由现有 Provider 测试负责。

## 验证

| 层级 | 结果 |
| --- | --- |
| L1 可复用合同定向测试 | net6.0、net8.0 各 9/9；NPOI、ClosedXML 各执行四项合同，另有缺失服务入口自检 |
| L2 Provider Contract `Category!=Large` | net6.0、net8.0 各 291/291 |
| L3 独立 NuGet 包消费者 | 双 TFM 还原、构建与运行通过，均输出 `third-party-public-only-provider-ok` |
| L3 Docs | net8.0 10/10 |
| L0 Solution Release build | 0 错误；4 个既有 Microsoft.Bcl.Memory net6.0 警告 |

模板首次运行时，测试数据的动态 Key 与标题仅大小写不同，按现有唯一性规则被正确拒绝；改为不同 Key 后通过。名称断言首次只接受单个地址，ClosedXML 实际合法写法为 `Contract!$A$4:$A$4`；按同一单元格范围修正断言后双 Provider 通过。这两项均为测试模板修正，生产代码未变。

本轮 Change Impact 为测试辅助源码、Provider Contract、包消费者和文档。生产实现和包内容未变，因此复用上一阶段 NPOI/ClosedXML 职责测试、双 TFM API 快照与常规全量测试证据；未重跑无关性能矩阵。Consumer 从当前 NuGet 包测试公开接口，模板源码由项目文件链接，未引用 Testing 工程程序集。

`git diff --check` 无空白错误。本轮 C# 文件为严格 UTF-8 BOM+LF，csproj 为 UTF-8 CRLF，Markdown 为无 BOM+LF，均保留末尾换行。API baseline 和 ProfileFixtures.xml 的既存行尾规范化提示未在本阶段修改。未改动 Utopa 仓库，未提交、推送、创建 PR、发布或修改版本号。

## TODO 分类

- Completed：四项可复用合同、官方 Provider 直接执行、双 TFM 包消费者、文档和追溯。
- Open Actionable：无。
- Blocked Approval / Blocked External：无。
- Not Applicable：生产 API 差异审批、API baseline 更新和性能 before/after。
- Accepted Limitations：模板验证 XLSX 公共行为；XLS/HSSF、格式专属资源及失败矩阵由 Provider 自身测试。
- Deferred：真正发布独立 NuGet 测试包需要单独确定版本和发布流程；Roslyn 地址/动态 Key 分析器为下一独立阶段。
- Next Action：本阶段停止；后续先评估 Roslyn 分析器的构建接入与误报边界。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
REVIEW_STATUS: NARROW_SCOPE_CHECKED
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
