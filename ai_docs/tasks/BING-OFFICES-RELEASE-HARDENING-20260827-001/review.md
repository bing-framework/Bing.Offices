<!-- AI_REVIEW_STATUS: NEEDS_FIX -->
AI_TASK_ID: BING-OFFICES-RELEASE-HARDENING-20260827-001
AI_REVIEWED_AT: 2026-08-27T17:52:48.7212233+08:00

# Round 4 独立复审报告

## 验收摘要

最终结论：`NEEDS_FIX`。

Round 3 对 MUST_FIX 有实质进展：FIX-005 已解决上一轮指出的探针 mode 无效、重复小文件和异常成功退出问题；FIX-001/002 已增加内部 IO 边界、故障注入和真实 Windows 文件系统场景，但仍未完整满足上一轮明确要求的双故障与权限/占用矩阵。FIX-003、FIX-004、FIX-006、FIX-007 未纳入 Round 3 `fixScope=must`，当前仍未解决。根据 Review 协议，存在未解决 MUST_FIX 或 SHOULD_FIX 时不能判定 PASS/PASS_WITH_ISSUES。

当前版本继续保持用户指定的 `2.0.0`。本 Review 只修改 `review.md`，未修改业务代码、测试、版本、`plan.md` 或 `execution.md`，未执行 Git 写操作或发布操作。

## Review 边界

- 当前 Diff 仍集中于 metadata、原子提交、Failure Workbook、CSV/Excel 资源限制、API 门禁和对应测试，属于本任务范围。
- Round 3 新增 Core internal IO friend assembly、独立资源探针 workload 和 Windows 文件故障 Integration，没有扩散公共 API。
- `git diff --check` 通过，仅有 runtime JSON 的 LF/CRLF 提示。
- Phase 3～6、Benchmark、办公客户端互操作和最终发布材料仍未完成。

## 上一轮 FIX 复验

| FIX | 复审状态 | 证据与结论 |
| --- | --- | --- |
| FIX-001 | PARTIAL | `IAtomicFileSystem` 已允许确定性注入，写入失败+删除失败会保留 Excel 主异常与 `TemporaryCleanupException`；Windows 目标锁定场景通过并保留旧内容、清理 staging。但没有提交失败+删除失败组合，也没有 CSV 双故障键、无新目标及残留路径的完整断言。 |
| FIX-002 | PARTIAL | 目录创建、临时文件创建、destination 复制失败 Unit 均通过；真实目录路径与普通文件冲突 Integration 通过。上一轮明确要求的安全 Windows 权限/文件占用场景仍不存在，创建失败测试也未断言 destination 未污染及无临时残留。 |
| FIX-003 | NOT_RESOLVED | XLS/XLSX 显式 metadata override 已覆盖；默认 preserve 仍只有 XLSX，未发现顺序隔离、64 并发导出重开或 Docs package consumer 对 `Metadata(...)` 的调用。 |
| FIX-004 | NOT_RESOLVED | `NpoiExcelImporter` 先构造 `existingSheets`，进入导入循环后仍再次调用 `ResolveSheetIndex`；一次解析结果没有贯穿 plan/materialization，完整 selector 矩阵仍不足。 |
| FIX-005 | RESOLVED | 七个 mode 使用不同受控 XLSX/XLS workload；`dom-limit` 实际设置 `MaxRows=100` 并返回 `resource-limit`；输出规模、结构、耗时和峰值工作集指标；未处理异常返回退出码 1。独立复验通过。 |
| FIX-006 | NOT_RESOLVED | 2.0.0 package-only consumer 仍通过 9/9，但 `package-consumer.md` 仍使用 `F:\Data\NuGetPackages` 预热缓存并将本地 Bing 包目录作为唯一 source，未准确说明第三方依赖来源。 |
| FIX-007 | NOT_RESOLVED | Phase 3～6、Benchmark、互操作和最终发布材料未实施，也没有用户批准的具名 release waiver。 |

## 计划验收矩阵

| 计划项 | 结论 | 实际证据 |
| --- | --- | --- |
| RH27-000 基线与追溯 | PARTIAL | 有 API、包和测试证据；缺完整生产符号追溯、benchmark 原始产物和发布清单。 |
| RH27-101 selector | PARTIAL | 冲突检测成立；单次解析复用和完整 XLS/XLSX 矩阵未完成。 |
| RH27-102 metadata | PARTIAL | 请求快照、XLS/XLSX override 成立；默认 XLS preserve、并发隔离和 consumer 缺失。 |
| RH27-103 File/Stream | PARTIAL | 原子主链和真实目标锁定成立；提交+清理双故障与 CSV 键矩阵不完整。 |
| RH27-104 Failure Workbook | PARTIAL | 创建/复制/删除故障注入成立；真实权限/占用和残留矩阵不完整。 |
| RH27-105 资源限制 | PASS | CSV 公开错误合同及 Excel 独立进程 workload、超限状态、指标和失败退出成立；未夸大 DOM 前置保护。 |
| RH27-106 异常/释放 | PARTIAL | CSV 枚举器与资源错误路径有覆盖；完整反射异常矩阵未完成。 |
| RH27-107 ADR | PASS | DOM、公式缓存和取消边界表述准确。 |
| Phase 1 门禁 | FAIL | FIX-001/002 仍有 MUST_FIX 证据缺口。 |
| Phase 2 API/包 | PASS | API gate 与当前 2.0.0 package-only consumer 主链成立。 |
| Phase 3 职责重构 | FAIL | 未执行。 |
| Phase 4 测试体系 | PARTIAL | net6/net8 回归全绿；P0 文件双故障和权限矩阵不完整。 |
| Phase 5 Benchmark | FAIL | 无本轮 BenchmarkDotNet 结果。 |
| Phase 6 文档/发布 | PARTIAL | API/迁移文档存在；互操作、release checklist 和最终报告缺失。 |

## 功能与真实接入 Review

### 原子文件提交

Excel/CSV File API 继续共用 `AtomicFileCommitter`。新增 internal `IAtomicFileSystem` 没有进入公共 API，默认实现仍使用同目录随机 `CreateNew`、`Flush(true)`、`File.Replace/File.Move`。写入失败+删除失败测试证明主异常优先和结构化清理诊断；Windows 目标锁定 Integration 证明旧目标内容保留且 staging 被清理。剩余缺口是上一轮要求的提交失败+删除失败及 Excel/CSV 双格式诊断矩阵。

### Failure Workbook

创建目录、创建临时文件和复制 destination 的稳定错误分类已有直接 Unit；路径与普通文件冲突的真实磁盘场景会进入 Failure Workbook 并返回稳定目录创建错误。该场景证明真实文件系统异常，但不是 Windows 权限或目标占用故障；创建失败测试也没有完整检查 destination 长度与 temporary residue。

### 独立资源探针

测试按 mode 生成不同 workload：DOM 250x4、shared strings 20x20、styles 20x20/80+、drawings 6 张、XLS OLE 和小型 ZIP 基线。探针读取实际 workbook 指标，`dom-limit` 使用请求资源限制，父进程断言状态与 workload 数值；异常返回非零。该实现解决上一轮 FIX-005 的具体问题。现有规模属于受控正确性证据，不应包装为 100K 性能预算或压缩炸弹防护。

## API、架构与维护性 Review

- `IAtomicFileSystem` 和 friend assembly 仅用于内部测试边界，没有新增公共 IO 抽象。
- 当前生产 IVT 从零增加为 Core 对 Unit/Integration 两个测试程序集；用途局限于原子提交 internal 合同，需在最终追溯矩阵中记录。
- Selector 仍重复解析，Phase 3 要求的 resolver/编排拆分未完成。
- 当前版本和 PackageReference 保持 `2.0.0`，本轮没有重新打开版本号争议。

## 性能与资源 Review

- 独立进程现在能证明不同 workload、超限状态和指标采集，不再 catch-all 成功。
- workload 规模较小，没有基线/压力 Peak Working Set 阈值，也没有 BenchmarkDotNet 或 1K～100K 正式预算；这些属于 FIX-007/Phase 5，而非重新打开已解决的 FIX-005。
- NPOI 仍为 DOM 管线，当前文档没有宣称能在 WorkbookFactory 前阻止全部解压后峰值。

## 测试 Review

| 验证 | 独立结果 |
| --- | --- |
| Round 3 FIX 定向 Unit/独立进程 | PASS，6/6 |
| Windows 文件故障定向 Integration | PASS，2/2 |
| net8 Unit 全量 | PASS，354/354 |
| net6 Unit 全量 | PASS，354/354 |
| net8 Integration 全量 | PASS，14/14 |
| net6 Integration 全量 | PASS，14/14 |
| Docs 2.0.0 package-only consumer | PASS，9/9 |
| 编辑器错误检查 | PASS，相关文件无错误 |
| `git diff --check` | PASS，仅 runtime JSON 换行提示 |

现有测试没有回归。测试缺口集中于尚未满足的 FIX-001/002 矩阵，以及此前延期的 metadata、selector 和发布门禁。

## 文档 Review

- `execution.md` 已如实记录 Round 3 `fixScope=must` 与全量回归结果。
- `package-consumer.md` 仍未说明 `F:\Data\NuGetPackages` 必须预热第三方包；当前命令不能证明空缓存且只有本地 Bing 源即可恢复。
- Benchmark、互操作、release checklist 和最终 Go/No-Go 报告仍缺失。

## 问题分级

### BLOCKER

无外部 BLOCKER。

### HIGH

1. FIX-001 的提交失败+删除失败、CSV 清理诊断键和残留状态矩阵不完整。
2. FIX-002 缺安全的 Windows 权限/占用真实故障及创建失败 destination/residue 断言。

### MEDIUM

1. metadata 默认 XLS preserve、顺序/并发隔离和 package consumer 调用仍缺失。
2. selector 解析结果未贯穿 plan/materialization，完整矩阵未完成。
3. package consumer 文档未准确说明第三方依赖恢复前提。
4. Phase 3～6、Benchmark、互操作和最终发布材料未完成且无具名 waiver。

### LOW

1. 既有弃用与 xUnit analyzer warnings 较多，本次未发现其导致行为失败。

## FIX 清单

### FIX-001

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH27-103、RH27-401
- 涉及文件/符号：`AtomicFileCommitter`、`FailingAtomicFileSystem`、Excel/CSV File API tests
- 问题：internal IO 边界和写入+清理双故障已建立，但上一轮要求的提交+清理双故障及双格式诊断合同未闭环。
- 证据：现有双故障测试只设置 `FailWrite=true, FailDelete=true` 并断言 Excel 键；提交失败测试只设置 `FailMove=true`，删除成功；没有 CSV `TemporaryCleanupException`、无新目标或残留路径断言。
- 影响：不可逆提交阶段与清理同时失败时，主异常优先级和 Excel/CSV 一致性仍缺发布证据。
- 修复目标：完成统一 committer 的最小双故障矩阵，不扩大生产抽象。
- 明确修复要求：增加提交失败+删除失败测试，断言主提交异常、对应 Excel/CSV 清理异常键、旧目标保留/新目标不存在、删除被调用及 staging 残留可诊断；覆盖 Move 和 Replace 中至少实际使用的失败路径。
- 修复后的验证方式：net6/net8 定向 Unit；Windows 目标锁定 Integration；net6/net8 全量回归。

### FIX-002

- 严重程度：`HIGH`
- 处理要求：`MUST_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH27-104、RH27-401、RH27-403
- 涉及文件/符号：`NpoiFailureWorkbookWriter`、`IFailureWorkbookFileSystem`、Failure Workbook tests
- 问题：受控创建/复制故障和真实路径类型冲突已覆盖，但上一轮明确要求的 Windows 权限/文件占用证据及完整残留断言仍缺失。
- 证据：Integration 只把普通文件作为 `TemporaryDirectory`；没有无权限目录或目标占用场景。目录/文件创建失败 Unit 只断言消息与 inner exception，未断言 destination 未污染、删除调用和临时目录残留。
- 影响：真实权限/占用环境中的错误分类、数据不泄漏和残留合同仍未达到 P0 故障矩阵。
- 修复目标：以安全、可清理、不填满磁盘的 Windows 场景闭环创建/复制失败，并补全 Unit 状态断言。
- 明确修复要求：增加 Windows 临时目录权限或文件占用 Integration；若 ACL 在 CI 不稳定，可使用真实独占锁定文件路径并明确证明命中 CreateFile/Copy 路径。Unit 必须断言 destination 长度不变、创建失败无临时残留；异常/诊断不得包含工作簿业务内容。
- 修复后的验证方式：故障注入 Unit + Windows Integration；net6/net8 回归。

### FIX-003

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH27-102
- 涉及文件/符号：metadata tests、Docs package consumer
- 问题：请求级 metadata 的格式与并发验收矩阵未完成。
- 证据：默认 preserve 测试仅覆盖 XLSX；未发现顺序隔离、64 并发导出重开或 Docs consumer `Metadata(...)` 调用。
- 影响：HSSF 默认路径和请求隔离缺少发布证据。
- 修复目标：完成请求级 metadata 的格式与并发合同验证。
- 明确修复要求：增加 XLS 默认 preserve、顺序隔离和 64 个不同 metadata 并发导出重开测试；Docs 2.0.0 package consumer 编译并调用 metadata API。
- 修复后的验证方式：net6/net8 Unit/Integration；2.0.0 package consumer。

### FIX-004

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH27-101、RH27-301
- 涉及文件/符号：`NpoiExcelImporter.Import`、`NpoiImportPlanBuilder`、selector tests
- 问题：selector 仍被解析两次，未实现一次解析结果贯穿主链。
- 证据：同一导入方法先构造 `existingSheets`，随后在 `foreach (request.Sheets)` 中再次调用 `ResolveSheetIndex`。
- 影响：comparer、缺失/隐藏 Sheet 或未来 resolver 变化时可能出现 plan 与 materialization 漂移。
- 修复目标：复用确定的 resolved physical sheet 结果。
- 明确修复要求：建立 request -> physical index/sheet 结果并传入 plan、导入和 failure mapping；补 comparer、越界、缺失、隐藏、合法混合、相同模型不同 mapping，以及 XLS/XLSX 场景。
- 修复后的验证方式：selector 定向 Unit/Integration；net6/net8 全量回归。

### FIX-006

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：RH27-404、RH27-603
- 涉及文件/符号：`package-consumer.md`、Docs restore 命令
- 问题：2.0.0 package consumer 主链成立，但验证文档没有准确说明第三方依赖需要预热缓存或额外可信源。
- 证据：命令设置 `NUGET_PACKAGES=F:\Data\NuGetPackages`，同时只指定本地 Bing 包 source；空缓存不能从该目录恢复 NPOI、CsvHelper、xUnit 等依赖。
- 影响：其他环境按文档复现时会误判为包不可消费。
- 修复目标：让 package consumer 验证前提和命令可复现。
- 明确修复要求：明确本地目录只提供 Bing 2.0.0 包；分别给出预热离线缓存和显式可信第三方源的恢复方式，不声称空缓存仅本地 Bing 源即可恢复。
- 修复后的验证方式：按文档命令在声明的环境恢复；assets 精确为 2.0.0，无 NU1601；Docs 9/9。

### FIX-007

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 当前状态：`OPEN`
- 对应计划项：Phase 3～6、RH27-604
- 涉及文件/符号：职责拆分、测试矩阵、Benchmark、互操作、发布报告
- 问题：原计划剩余门禁未完成，也没有用户批准的 release waiver。
- 证据：execution 明确未重跑 Benchmark、办公客户端互操作和全 TFM；没有最终发布报告。
- 影响：任务仍不能达到原 plan 的 COMPLETE 或发布 Go。
- 修复目标：完成剩余门禁，或取得明确、具名、带风险的用户 waiver。
- 明确修复要求：按 Phase 3～6 补最小必要证据；性能绑定当前 diff；客户端不可用标 NOT_VERIFIABLE；不得把 Unit 全绿包装为发布 Go。
- 修复后的验证方式：build/test/docs/pack/consumer/benchmark/互操作矩阵和最终独立 Review，或核验用户批准 waiver。

## 已解决项

### FIX-005

- 复审状态：`RESOLVED`
- 证据：独立 mode workload、真实 `dom-limit`、结构化指标和异常非零退出均进入实际测试；定向 6/6 及全量 net6/net8 回归通过。
- 说明：受控资源正确性证据不替代 Phase 5 正式性能预算，后者继续由 FIX-007 跟踪，不重复打开 FIX-005。

## 未完成与偏离

- 当前 2.0.0 是用户明确版本决策，继续标记 `DEVIATED_OK`。
- Round 3 按用户要求只处理 MUST_FIX；FIX-003/004/006/007 未处理不是回归，但仍是 Review 协议中的开放 SHOULD_FIX。
- `execution.md` 的 `COMPLETED` 表示 Round 3 `fixScope=must` 执行结束，不代表原计划 COMPLETE 或 Reviewer PASS。

## 回归与兼容风险

- 原子提交/Failure Workbook 双故障仍可能遗留包含业务数据的 staging，完整诊断矩阵尚未证明。
- Selector 重复解析和 metadata 并发证据缺失，未来修改可能造成 plan/materialization 漂移或跨请求污染。
- NPOI DOM 峰值需要部署隔离和 Phase 5 预算，当前受控探针不等于压缩炸弹防护。

## 最终验收 Checklist

- [x] 读取 plan、最新 execution、旧 review、当前源码、测试与 Git Diff
- [x] 优先逐项复验 FIX-001/002/005
- [x] 独立运行定向 Unit、独立进程与 Windows Integration
- [x] 独立运行 net8/net6 Unit 与 Integration 全量回归
- [x] 复验 Docs 2.0.0 package-only consumer
- [x] 检查错误与 `git diff --check`
- [x] FIX-005 独立进程 workload、超限和失败退出闭环
- [ ] FIX-001 提交+清理双故障及双格式矩阵闭环
- [ ] FIX-002 Windows 权限/占用与残留矩阵闭环
- [ ] FIX-003/004/006/007 完成或取得具名 waiver

最终状态保持 `NEEDS_FIX`。下一步应根据本报告执行 Review Fix；Reviewer 不自动修复。
