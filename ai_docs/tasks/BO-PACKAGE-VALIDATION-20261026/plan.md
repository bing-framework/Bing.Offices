# 新 API 包消费与候选验证方案

Task ID: BO-PACKAGE-VALIDATION-20261026

## 授权与现状

用户在上一轮复查后要求“继续实现相关功能”，授权完成当时列出的本地验证和 CI 接入缺口。先形成此方案，再按 execute-plan 实施；无需重复执行确认。旧 BO-API-USABILITY-20261025 已验收的计划及 Review 作为历史输入保留。

- 已完成：Core 注册、六个单表入口、统一导入状态和两个 P2 修复；双框架 66 次复查通过，完整套件历史记录 1,999 次通过。
- 部分完成：tests/Bing.Offices.Consumer.Net6/Net8 是真实 PackageReference 消费项目，但只验证旧 Workbook 入口；没有仅引用 Core 的独立消费项目。快速开始测试仍使用源码项目及宿主引用。
- 待补齐：build/ApiSnapshot/Program.cs 在比较时从批准基线读取候选身份，导致新源码无法使用历史 API 基线验证自己的新候选；CI 已有 pack 和消费者流程，尚未调用完整快照身份检查。
- 外部待验证：匹配当前候选的远端 CI 和生产资源/并发部署证据。本轮完成配置、本地执行与可归档材料，不推送、发布或触发部署，不以本地结果代替远端成功。

## 设计边界

不改变生产公共 API、默认行为、包版本、依赖版本、TFM、DOM 并发和资源策略。复用既有消费项目、API 规范化及身份算法，避免另一套快照规则。批准的 build/api-snapshot-baseline.json 不修改。

独立消费者由脚本复制必要项目/源码到仓库外临时目录，只从隔离本地包源恢复 Bing.Offices.*；每个消费项目使用独立空包缓存。第三方依赖仍从 NuGet 恢复。Core 项目只引用 Bing.Offices.Core 与 DI 实现，并检查恢复图没有 NPOI。

新增 --candidate 参数只选择当前身份清单；API 成员/审批始终来自 --baseline。未传该参数时保留旧严格比较路径。capture 输出不自动成为批准基线。输出保存实际候选和成员差异，失败退出非零。

## Phase 1 / PKG-01：独立包消费者

- 目标：在 net6/net8 实际安装包后验证 AddBingOfficesCore、ImportSheet 和状态摘要。
- 已确认修改范围：现有两个 Consumer 项目；新增 Core-only 双 TFM 消费项目及必要共享验证源码。
- 步骤：保留旧入口验证；增加六个单表入口的 XLS/XLSX 成功及结构化失败验证；检查完整数据、源行、IsSuccess/IsComplete/StopReason。Core-only 检查 CSV 同步/异步完整输出、完整错误和状态，以及无 Excel Provider 注册。
- 依赖：无；窄范围消费者源码委派 luna_worker，主代理负责脚本/集成。
- 验证：最终脚本从本轮打包产物在隔离目录构建和运行两个旧消费者与 Core 的两个 TFM，共四个执行目标。
- 风险：偷用工作区项目引用、NPOI 隐含依赖、只检查行数。验收要求包源锁定本地，Core 恢复图无 NPOI，全部内容/错误断言通过。

## Phase 2 / PKG-02：候选身份与批准 API 分离

- 目标：保留批准 API 合同，同时验证当前源文件、程序集和包的逻辑身份。
- 已确认修改范围：build/ApiSnapshot/Program.cs；既有 CandidateIdentityContractTest.cs 如需要；候选脚本的 CLI 反例验证。
- 步骤：加入独立 --candidate 清单；校验清单格式及身份，API 比较始终使用原 --baseline；比较输出保留候选身份。保留旧调用兼容性和历史审批文件校验。
- 依赖：无。
- 验证：现有 identity-self-test；真实候选 capture/compare；故意损坏候选身份、使用缺失清单、改变批准 API 的临时副本，均非零退出并包含对应诊断。
- 风险：把 capture 当 API 审批、让候选参数绕过基线。验收要求原基线零修改，当前正常候选通过，错误 API 即使候选合法也拒绝。

## Phase 3 / PKG-03：可重复脚本与 CI

- 目标：本地与 CI 使用同一条 pack → 身份捕获 → 隔离消费 → 严格比较链路。
- 已确认修改范围：新增 build/Verify-PackageCandidate.ps1；.github/workflows/ci.yml；新增 docs/excel/package-validation.md。
- 步骤：生成专属输出目录；打包三个生产项目并检查 nupkg/snupkg、lib/XML；在临时目录生成 UTF-8 NuGet.Config，配置包源映射并执行四个消费目标；保存 assets、构建/运行日志、API 候选、差异和摘要；finally 清理自建临时目录。替换 CI 现有重复 pack/消费步骤，并用 always 上传证据。
- 依赖：PKG-01/02。
- 验证：Windows PowerShell 实际运行；PowerShell 语法、YAML 结构及 git diff --check；CI 使用 pwsh 执行同脚本。以现有 NuGet 配置和真实工具命令为准。
- 风险：同版本缓存误命中、旧输出混入、脚本吞非零退出、远端日志不可审计。验收要求隔离输出/缓存、清晰失败及完整证据；输出目录非空时拒绝覆盖。

## Phase 4 / PKG-04：收口与证据

- 目标：记录实现与外部门禁的真实状态。
- 修改范围：本任务 execution.md、symbol-test-map.md，验证使用文档。
- 验证：完整脚本、快照自检和反例；复用未变更生产范围的既有 Unit/Integration/Docs 证据，不重跑性能或生产容量矩阵。
- 验收：三个本地缺口有实现和可复现证据；外部 CI/部署仍清楚标记未验证。没有 Open Actionable 后停止，不为外部状态循环修复。

## 影响分析与命令

ChangedProjects=Consumer/CoreConsumer/ApiSnapshot；ChangedPublicContracts=生产 API 无；ChangedRuntimePaths=构建验证工具；ChangedProviders/TFMs=无；ChangedBuildPackaging=验证流程；ChangedBenchmarkHarness=无；RiskLevel=MEDIUM。

工具命令基于现有项目：dotnet pack 三个 src 项目；dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release -- --identity-self-test true；dotnet restore/build/run 现有消费者。新增统一命令：pwsh -File build/Verify-PackageCandidate.ps1 -OutputDirectory <新的证据目录>。

先完成工具/消费者构建与自检，再运行一次完整候选验证；仅失败相关范围重试。历史报告、批准基线和用户已有未提交改动均保留；不 commit/push/PR、不发布包、不发送通知。按 task-state start 注册，最终 task-finish --no-notify 收口。
