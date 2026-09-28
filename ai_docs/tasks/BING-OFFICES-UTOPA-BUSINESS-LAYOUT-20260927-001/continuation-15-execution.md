# Continuation 15 执行记录

## 结果

新增独立的双 TFM `Bing.Offices.EntityProviderContracts` 测试包，直接编译现有 `EntityLayoutProviderContractSuite`，保留既有命名空间和六项合同入口。包的 net6.0、net8.0 nuspec 依赖均仅为 `Bing.Offices.Abstractions 2.0.0`，不引入 xUnit 或官方 Provider。第三方消费者删除源码链接，改从本地包加载合同，并执行全部六项合同。

CI 在生产包验证后将合同包单独打入 `artifacts/packages-contracts`，检查两个程序集及依赖，再使用专用 NuGet 配置还原第三方消费者。原八个生产包数量门禁保持不变；包未发布，版本号未修改。

## Change Impact

- ChangedProjects：新增合同包工程，第三方包消费者引用方式和合同调用，CI、文档与任务追溯。
- ChangedPublicContracts：新增测试用途的独立包；原八个生产程序集成员与签名不变，既有 API snapshot 不受影响。
- ChangedRuntimePaths / ChangedProviders：无；上一阶段 Provider 测试证据按未变源码范围复用。
- RiskLevel：MEDIUM；重点是包依赖、双 TFM 资产及消费者脱离源码链接后的实际运行。

## 验证

| 检查 | 结果 |
| --- | --- |
| 合同包还原与 Release 打包 | 成功；包含 net6.0、net8.0 DLL 和 XML 文档 |
| nuspec | 两个 TFM 均仅依赖 `Bing.Offices.Abstractions 2.0.0` |
| 第三方包消费者 | 隔离本地缓存还原、Release 构建通过；net6.0、net8.0 均输出 `third-party-public-only-provider-ok`，六项合同全部执行 |
| 专用 NuGet 配置 | 生产/合同源路径解析正确，包源映射下使用缓存还原通过 |
| Provider 合同定向测试 | net6.0、net8.0 各 13/13，通过 |
| Docs 原文围栏测试 | net8.0 10/10，通过 |

本地隔离缓存首轮还原因源映射排除了命令行追加的离线依赖源而失败；去掉映射后，命令行相对 `--source` 路径又被 NuGet 按消费者工程目录解释。使用绝对源路径完成首次还原后，恢复专用配置的精确包源映射，并在已缓存依赖上验证通过。消费者 `project.assets.json` 的两个 TFM 均指向合同包 DLL，工程文件没有源码链接。外部 CI 尚未在本地执行；本地复现了其包结构和消费者流程。

`git diff --check` 无空白错误；本阶段涉及的十个文本文件通过严格 UTF-8、BOM、行尾和末尾换行检查。既存 API baseline 与 ProfileFixtures.xml 的 Git 行尾提示未在本阶段处理。未修改 Utopa 仓库，未提交、推送、创建 PR 或发布包。

## TODO 分类

- Completed：独立合同测试包、双 TFM 包内容及依赖验证、第三方包消费、CI 验证步骤和接入文档。
- Open Actionable：无本阶段未完成项。
- Blocked Approval：无。
- Blocked External：远端 CI 实际运行结果尚不可得；本地构建与消费证据不冒充外部 CI。
- Not Applicable：原生产程序集 API 差异与性能 before/after；本阶段未修改对应源码。
- Accepted Limitations：合同当前限定 XLSX Entity Layout；NPOI XLS/HSSF 专属行为由 Provider 测试覆盖。
- Verified Boundaries：无。
- Deferred：Footer 公式/签字区；旧版迁移分析器；有性能证据后评估 Source Generator 和显式 DOM/Streaming 策略。
- Next Action：阶段工作停止；后续独立设计 Footer 业务表达能力。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
RELEASE_STATUS: NOT_EVALUATED
OPEN_ACTIONABLE: 0
