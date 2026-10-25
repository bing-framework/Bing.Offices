<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BO-API-USABILITY-20261025
AI_EXECUTION_FINISHED_AT: 2026-10-25T12:33:04Z

# 实施执行报告

## 当前结论（Review 修复 Round 1）

已完成本轮 recommended 范围的 FIX-001、FIX-002，新增 32 组职责测试。当前双 TFM Unit/Integration 与 Docs 共 1,999 项通过，失败和跳过均为 0；公开成员契约与原快照一致。原 review.md 保留不改，本文的 COMPLETED 表示修复执行完成，不代替正式独立 Review 的状态。

IMPLEMENTATION_STATUS: PASS（本轮修复执行与验证）
TEST_STATUS: PASS
OPEN_ACTIONABLE: 0（本轮执行范围）
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL
RELEASE_STATUS: NOT_VERIFIED

本轮证据和影响分析见文末“Review 修复记录”。以下初次实施内容保留为历史记录；原包身份验证不作为本轮修改后的候选身份验证。

## 初次实施结论（历史）

已按 [plan.md](plan.md) 完成六项易用性发现及配套文档、契约维护。最终 1,935 项测试全部通过，失败 0、跳过 0；双 TFM API 与候选身份比较通过。新增公共面为兼容增量，未删除、重命名或改变旧签名。

IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS
OPEN_ACTIONABLE: 0
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL（未执行远端 CI 或生产部署验证）
RELEASE_STATUS: NOT_VERIFIED（本任务不发布）
PERFORMANCE_EVIDENCE_STATUS: NOT_APPLICABLE（本次无性能声明）
RESOURCE_EVIDENCE_STATUS: REUSED_UNCHANGED_SCOPE（文件提交、DOM 并发和 IO 路径未改）

## 任务信息

- 用户授权：先做解决方案，再实现已发现问题；计划落盘后继续实施，未要求重复批准。
- 基线 HEAD：`7496bd40a29213e0d0ba03836943ffb02d5e7bc8`，加此前任务已有未提交修改。
- 开始：2026-10-25T06:44:59.722Z；执行器：Codex；环境：Windows/PowerShell。
- 窄范围任务委派 luna_worker：默认值职责测试、文档原文校验、快速开始；主代理负责 API 设计、生产实现、集成与最终验证。

## 计划执行情况

| 项目 | 状态 | 已完成事项与验收 |
| --- | --- | --- |
| U-01 正则示例 | CLOSED | 修正 verbatim 正则转义；Markdown 模型真实导入 ORD-123 成功、BAD 返回完整错误 |
| U-02 模板所有权 | CLOSED | XML 与文档明确 leaveOpen=false 默认接管、true 保留；复用并运行模板成功/失败/取消职责测试 |
| U-03 默认语义 | CLOSED | 明确 StopOnFirstFailure 按行；100 列默认上限、101 列拒绝、显式扩容及多行继续都有 XLS/XLSX × Sync/Async 证据 |
| U-04 Core 接入 | CLOSED | AddBingOfficesCore；Core-only CSV、默认校验、幂等和自定义服务优先；原 NPOI 注册复用它 |
| U-05 单表导入 | CLOSED | 六个 Stream/File/Bytes Sync/Async 入口，不要求用户根模型；保留全局错误、成功实体和源行；取消/异常不吞掉 |
| U-06 结果摘要 | CLOSED | CSV/Workbook/单表提供 IsComplete/StopReason，CSV 增加 IsSuccess；修复 Sheet 边界耗尽行限额后静默跳过 |
| U-07 文档与契约 | CLOSED | 完整 CSV/Excel 程序、导航、迁移、双 TFM 快照与成员追溯均完成 |

## 修改文件

生产新增：Abstractions 的 `Imports/ImportStopReason.cs`、`Imports/ExcelSingleSheetImportResult.cs`；Core 的 `Extensions/BingOfficesServiceCollectionExtensions.cs`、`Extensions/ExcelSingleSheetImportExtensions.cs`。

生产修改：`CsvImportResult.cs`、`ExcelWorkbookImportResult.cs` 增量状态；NPOI `Extensions.Service.cs` 注册委托及 `NpoiExcelImporter.cs` Sheet 边界诊断。`ExcelExport.cs`、`ExcelImport.cs`、`ValidateMode.cs` 仅本任务所述 XML 说明变化。

测试新增：Unit 的 `CoreServiceRegistrationTest.cs`、`ExcelSingleSheetImportTest.cs`、`ImportResultStatusTest.cs`、`UsabilityDefaultsTest.cs`；Docs 的 `QuickStartConsumerTest.cs`。修改 `DocsConsumerTest.cs` 的正则真实行为断言、`PublicApiContractTest.cs` 的类型分类、`PublicExtensionCoverageTest.cs` 的完整签名数量和追溯。

文档与契约：README、docs/excel/{README,quick-start,import-validation,async-io,nuget-migration}.md、build/api-snapshot-baseline.json、本任务 plan/execution/api-changes/symbol-test-map。

工作区还有此前的 CI、表头优化、CSV 修复、benchmark 与其他任务文件，均保留，不将它们混记为本任务新实现。

## API / 数据 / 配置变化

- [api-changes.md](api-changes.md) 列出每个 TFM 的全部 31 条新增 canonical 记录：Abstractions 22、Core 9、NPOI 0；包含类型与枚举编译字段。无旧记录删除或签名变化。
- 原 Provider SPI、Workbook Request、IExcelImporter、旧构造器和结果字段全部保留；不增加依赖、不修改版本、不修改持久化数据格式。
- `IsComplete` 是结构拒绝/限制的摘要，不等于文件物理 EOF，也不等于所有行有效。旧 Workbook `IsSuccess` 仍只检查 Errors；手工构造空错误却截断的矛盾结果时要结合 IsComplete，文档已说明。
- 多原因优先级为 ResourceLimit → ErrorLimit → InvalidHeader → Unspecified；错误明细不丢失。显式配置范围结束算完整。
- 唯一既有运行行为补正：MaxRows 恰在 Sheet 边界耗尽且还有请求 Sheet 时返回 ResourceLimit；最后一张请求 Sheet 恰好额度不误报。
- 授权只覆盖计划内兼容增量及其强制快照维护，不作为历史或其他破坏性变化的审批。旧基线冻结于 `TestResults/usability/api/previous-baseline.json`，最终捕获于 `TestResults/usability/api-final`。

## 测试结果

| 范围 | 通过 / 总数 | 最终证据（TestResults/usability/final） |
| --- | --- | --- |
| Unit net6.0 | 924 / 924 | unit-net6-final.trx |
| Unit net8.0 | 924 / 924 | unit-net8.trx |
| Integration net6.0 | 38 / 38 | integration-net6.0.trx |
| Integration net8.0 | 38 / 38 | integration-net8.0.trx |
| Docs net8.0 | 11 / 11 | docs-net8.trx |
| 合计 | 1,935 / 1,935 | 失败、跳过均为 0 |

两份快速开始按 Markdown 原文分别编译为 ConsoleApplication 后执行，没有隐藏 using/变量/DTO；CSV 的编译引用明确排除 NPOI/Bing.Offices.Npoi。文档包含 .NET 8 Console、DI 实现包及源码/本地包版本说明。

单表六入口 XLS/XLSX 的成功、InvalidHeader、取消与 Provider 异常均直接验证；真实格式测试不以 Mock 替代。Core 默认实现有独立 CSV Sync/Async 测试。完整输出/结构化错误、源行、注册顺序、失败后复用均有职责断言。详细符号映射见 [symbol-test-map.md](symbol-test-map.md)。

扩展覆盖门禁最终 Core 22/22、NPOI 66/66，共 88/88。报告为 `extension-coverage-net6.md` 和 `extension-coverage-net8.md`；仍使用真实 IL call/callvirt 判定。

## Build / Typecheck / Lint / Format

- Release solution build：通过，0 warnings / 0 errors；生产双 TFM 构建通过。
- 快速开始定稿后的 Docs 构建/测试通过；消除了测试中忽略 Assert 返回 Task 引发的 CS4014，直接 await 该 Task。
- 三个本地包 pack 通过，仅用于快照身份验证，未发布。
- 最终 ApiSnapshot compare：net6.0 / net8.0 均 PASS；目录 `TestResults/usability/api-verified-final` 包含完整候选、双 TFM 快照与 api-diff.json。
- docgen docs-check：6 个源码示例、6 处文档内嵌示例一致，通过。
- git diff --check：通过；仅 Git 对既有 CRLF/LF 的提示，无空白错误。新增文件额外检查 UTF-8 与尾随空白。

最终 API 比较命令：

```powershell
dotnet run --project build/ApiSnapshot/ApiSnapshot.csproj -c Release --no-build -- --root output/release --baseline build/api-snapshot-baseline.json --dependencies tests/Bing.Offices.Tests/bin/Release --packages TestResults/usability/packages --repository . --output TestResults/usability/api-verified-final
```

## 计划偏差与修复经过

1. 初次完整 net6 测试为 923 成功、1 失败：扩展覆盖门禁仍固定 Core 15 / 总计 81。更新为真实新增后的 22 / 88，并让六入口在测试体内直接调用；未削弱门禁。原失败证据保留为 unit-net6.trx，最终通过为 unit-net6-final.trx。
2. 首次 API identity 比较时，README 正在由文档子任务修改；源码清单不一致。README 属于打包输入，等待定稿后重打三个本地包、重新捕获；断言 API assemblies 完全不变，仅更新最终 candidateIdentity。保留初始冻结旧基线与新增差异，最终比较通过。
3. 仓库旧 `tools/docgen/.venv` 的 lxml 扩展加载失败。未修改全局环境或生产依赖；使用 bundled Python，PYTHONPATH 指向 tools/docgen 与已有 TestResults/audit-fixes/docgen-python，直接 `-m docgen docs-check` 执行同一检查，成功。此项为本机环境边界。
4. 最后仅修改 Markdown、Docs 测试及快照身份，复用已完成的双 TFM Unit/Integration 证据，不因报告修改重跑全套。

## 基线问题与未完成事项

计划内无未完成事项，OPEN_ACTIONABLE=0。远端 CI、生产部署 gate 和更大规模生产资源证明属于已有外部范围，状态 BLOCKED_EXTERNAL；本地通过不替代这些证据。本轮没有发起部署或 NuGet 发布。

## 风险与回归关注点

不改变映射缓存键、计划缓存、IO staging、文件原子提交和 DOM 并发策略；因此不新建重型资源矩阵或性能 baseline，不作性能提升声明。新单表入口仍受 NPOI DOM 同步阶段约束。模板默认释放行为、列上限和按行校验策略没有改动。

## Reviewer 注意事项

先按 plan 的六项发现核对，再查 api-changes 与 symbol-test-map。测试总数是跨 TFM 执行次数之和；Docs 的 1 个快速开始测试内部独立运行两份完整程序。TestResults 为本地忽略产物，长期审计如需保留应由既有 CI artifact 流程归档；本任务文档保存命令、计数、签名及映射。

## Git 状态与收口

保留全部既有未提交改动；未自动 git commit、未自动 git push、未自动创建 PR、未发布包。最终使用 `task-finish.mjs BO-API-USABILITY-20261025 --no-notify` 收口，不发送外部通知。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX；Fix Scope：recommended。
- 输入：[review.md](review.md)；两个 SHOULD_FIX 均纳入，OPTIONAL 独立 NuGet 消费者未纳入本轮。
- 模式：REVIEW_FIX；开始时间：2026-10-25T12:24:20.443Z。
- 本轮生产范围仅 NPOI `NpoiExcelImporter.cs`；无 API、依赖、TFM、缓存、文件提交或 DOM 并发变更。
- 新增测试 `DefaultInputLimitTest.cs`、`SheetBoundaryDiagnosticTest.cs`；更新 async-io、import-validation、nuget-migration 及本任务 symbol-test-map。

#### FIX-002：默认输入预算

- 严重程度 MEDIUM/P2，处理要求 SHOULD_FIX，执行状态 COMPLETED。
- 根因：省略 ResourceLimits 时，nullable 成员访问向复制器传入 null，绕过文档规定的 128 MiB 默认值。
- 修复：Sync/Async 输入复制前先取得 `request.ResourceLimits ?? new ExcelResourceLimits()`，再读取 MaxInputBytes。显式对象的 MaxInputBytes=null 不做回填，继续保留主动关闭语义。
- 修复前证据：`default-limit-before.trx` 中 12 例通过 8、失败 4；失败均为 Workbook/单表 × Sync/Async 的默认超限场景。它们越过 128 MiB 后到达输入终点哨兵，证明缺省预算未接入；精确额度与显式关闭原行为正常。测试初次编译时误写异常代码枚举，已改用现有 ResourceLimitExceeded，未修改生产异常契约。
- 修复后：12 组实际默认阈值测试全部通过。默认 128 MiB+1 在复制时拒绝，Stage=Open；128 MiB 精确额度允许读取到终点哨兵；显式关闭允许读取超过默认阈值。哨兵只用于停止输入，避免解析大型无效 DOM，未用 Mock 替代 Provider。
- 另有 8 组真实 XLS/XLSX × Sync/Async × Workbook/单表文件测试：默认、显式精确额度、显式关闭均完整导入；额度比文件少一字节时返回完整资源异常。所有场景保留调用方流所有权。异步生成流拒绝同步 Read，真实文件流启用异步选项。
- 测试资源范围：合成输入实际供给 128 MiB 附近的字节，生产复制路径真实缓冲；测试集合禁止并行运行以限制同时占用。验证的是输入预算，不是 DOM 峰值、吞吐量或生产容量。
- 文档：新增默认输入配置说明、异常阶段及迁移提醒；依赖旧缺省无限制行为读取大文件的消费者须显式调整额度或关闭该项。

#### FIX-001：跨 Sheet 诊断表名

- 严重程度 MEDIUM/P2，处理要求 SHOULD_FIX，执行状态 COMPLETED。
- 根因：行限额诊断填写 selector 描述，未使用已解析物理表名。
- 修复：`resolvedSheet.Name ?? GetSelectorDescription(resolvedSheet.Request.Selector)`；存在时使用实际表名，缺失时保留回退。错误代码、0 行列、部分成功数据和只报告一次的行为不变。
- luna_worker 负责独立新增测试文件；主代理检查并集成。12 组 XLS/XLSX × Sync/Async × 索引/不同大小写名称/缺失名称回退测试通过。
- 断言覆盖全部结构化错误字段、成功实体和源行、单一错误，以及 AnnotatedOriginal 中 `_ImportErrors` 的完整八列表头与数据行。输入和失败输出流仍由调用方持有。

### Round 1 验证

最终证据目录：`TestResults/usability-fix-r1`，与初次实施及 Review 证据分开保存。

| 层级与范围 | 结果 | 证据 |
| --- | --- | --- |
| L1 新增 FIX 职责测试 net8 | 32/32 | targeted-net8.trx |
| L2 Unit net6 | 956/956 | unit-net6.trx |
| L2 Unit net8 | 956/956 | unit-net8.trx |
| L3 Integration net6 | 38/38 | integration-net6.0.trx |
| L3 Integration net8 | 38/38 | integration-net8.0.trx |
| L3 Docs net8 | 11/11 | docs-net8.trx |
| 当前去重套件合计 | 1,999/1,999 | 不重复计入 L1；跨 TFM 按实际执行次数累计 |

- Build：受影响测试项目及生产依赖在 Release 构建通过，无编译错误或警告；未因这三处 Provider 修复再构建无关 benchmark/整个 solution。
- API：双 TFM `PublicApi_AllReleaseAssemblies_ShouldMatchMemberSnapshot` 在完整 Unit 中通过；本轮没有公开成员变化，未更新批准基线或重打包。前轮 candidateIdentity 及 nupkg 属于历史候选，当前未声称通过新的包身份/发布门禁。
- Docs：沿用此前可用的 bundled Python 环境执行 `-m docgen docs-check`，6 处内嵌示例与源码一致。Markdown 代码消费者 11 项使用当前 Provider 重新构建执行。
- 格式：git diff --check 通过；两份新增测试严格按 UTF-8 读取并检查尾随空白，通过。既有 CRLF/LF 提示不构成文件损坏。
- 独立窄范围复核：luna_worker 未参与 FIX-002 的生产/测试编写，复查这部分后未发现剩余可执行缺陷；没有改写原 review.md 或冒充整轮正式 Review PASS。

### Round 1 影响分析与收口

ChangedProjects=Npoi + Unit；ChangedPublicContracts=无；ChangedRuntimePaths=输入默认预算和跨 Sheet 错误定位；ChangedProviders=NPOI；ChangedTFMs=无；ChangedBuildPackaging=无；RiskLevel=MEDIUM。先执行 L1，再执行受影响 Unit、Integration 和 Docs；没有必要运行新的重型 benchmark/容量矩阵。

- Completed：FIX-001、FIX-002，共 2/2。
- Open Actionable：无。
- Blocked Approval：无。
- Blocked External：远端 CI、生产部署资源/并发 gate；保持未验证，不以本地套件替代。
- Not Applicable：性能 before/after；本轮无性能提升声明。
- Accepted Limitations：NPOI DOM 同步及原始输入限额不代表 DOM 峰值限制。
- Verified Boundaries：实际 128 MiB 阈值，默认超出一字节拒绝，显式关闭可继续。
- Deferred：OPTIONAL 独立 NuGet 消费者 smoke，属于发布前验证增强，未发现依赖闭包缺陷。
- No-Progress Check：CHANGED，两个问题完成真实修复及回归证据。
- 下一步：停止当前修复，供独立 Reviewer 重新验收；不因原 review.md 保留 NEEDS_FIX 自动进入下一轮。
- 收口：`task-finish.mjs BO-API-USABILITY-20261025 --no-notify`。无 commit/push/PR，无外部消息或包发布。
