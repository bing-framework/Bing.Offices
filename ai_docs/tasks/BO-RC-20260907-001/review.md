<!-- AI_REVIEW_STATUS: PASS -->
AI_TASK_ID: BO-RC-20260907-001
AI_REVIEWED_AT: 2026-09-08T17:44:47.0785975+08:00

# 独立代码审查报告（Round 13 / Phase 9 修复复核）

## 审查结论

结论为 `PASS`。本轮仅复核 Round 11 遗留的 `FIX-030`、`FIX-031`、`FIX-032` 及 Round 12 修复证据，没有重新开展全仓分析。三项 `SHOULD_FIX` 均已完成，未发现新的 `MUST_FIX`、`SHOULD_FIX` 或 `OPTIONAL`。

本结论确认 Phase 9 修复达到当前本地验收要求，但不替代 GitHub push/PR CI 最终绿灯，也不解除外部生产入口进入 NPOI DOM 前共享并行度 1 门禁的发布条件。

## 修复复核

### FIX-030：IL 门禁真实调用判定

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 状态：`RESOLVED`
- 复核结果：`PublicExtensionCoverageTest` 仅在操作码为 `call` 或 `callvirt` 时解析被调用方法；普通测试只扫描测试本体，async 测试仅通过 `AsyncStateMachineAttribute` 精确加入状态机 `MoveNext`，不再按编译器生成名称扫描 lambda/本地函数。
- 反例证据：`CoverageReader_ShouldRejectMethodPointersAndUnexecutedLambdas` 同时验证真实直接调用可识别、仅方法组取址不可识别、未执行 lambda 不可识别。该反例与主门禁在四份 Phase 9 TRX 中均为 `Passed`。
- 结论：原 `ldftn`/`ldvirtftn` 和未执行 lambda 假覆盖风险已关闭。

### FIX-031：异常参数名与失败不变性

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 状态：`RESOLVED`
- 复核结果：行边界、合并区域移动、图片参数校验均以 HSSF/XSSF 双 Provider 参数化执行；测试精确断言 `deleteRowStartIndex`、`count`、`rowIndex`、`rowsCount`、`startRowIndex`、`endRowIndex`、`moveRowCount`、`moveColCount`、`picInfo`、`pictureData`、`pictureBytes`、`row`、`col`、`pictureType`、`sheet`。生产端重叠异常没有参数名，测试显式断言 `ParamName == null`。
- 不变性证据：行内容、行边界和合并区域在失败后保持原状；图片失败场景检查 workbook 图片集合与 sheet 图片信息不增加，已有 `IPictureData` 场景还校验失败前后图片总数一致。
- 结论：异常类型/参数合同及失败前无副作用要求已由双 Provider 职责测试固化。

### FIX-032：逐签名映射与原始测试证据

- 严重程度：`MEDIUM`
- 处理要求：`SHOULD_FIX`
- 状态：`RESOLVED`
- 复核结果：`public-extension-coverage.md` 解析得到 85 条数据行、85 个唯一完整签名，其中 Core 19、NPOI 66，测试项目均为 `Bing.Offices.Tests`，无重复签名。`symbol-test-map.md` 已指向该逐签名报告，并标明新的 Phase 9 TRX 替代旧 `BO-RC-continued-*` 证据。
- TRX 证据：net6/net8 完整 Unit 各 `716/716`，net6/net8 定向扩展各 `190/190`，均为 0 failed；四份 TRX 均包含并通过 `PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature` 和门禁反例测试。
- 结论：逐签名追溯、测试计数和原始结果已互相一致。

## Round 13 验收矩阵

| 验收项 | 结果 | 证据 |
| --- | --- | --- |
| FIX-030 门禁仅认真实调用 | `PASS` | `PublicExtensionCoverageTest.cs`：`call`/`callvirt`、`AsyncStateMachineAttribute`、方法组/未执行 lambda 反例；四份 TRX 双测试通过 |
| FIX-031 参数名与失败不变性 | `PASS` | `NpoiSheetExtensionsTest.cs`、`NpoiSheetPictureExtensionsTest.cs`：HSSF/XSSF 参数化、精确 `ParamName`、失败后状态断言 |
| FIX-032 85 行追溯映射 | `PASS` | `public-extension-coverage.md`：85 个唯一签名（19 + 66）；`symbol-test-map.md` 已建立入口 |
| 完整 Unit 原始结果 | `PASS` | net6/net8 各 `716/716`，0 failed |
| 定向扩展原始结果 | `PASS` | net6/net8 各 `190/190`，0 failed，均含主门禁和反例测试 |
| Round 12 既有验证 | `PASS` | `execution.md`：Integration 各 15/15、Docs 10/10、Release build 0 error（1 条既有 NU1900）、API compare 双 TFM PASS、`git diff --check` PASS |

## 发布条件

- `.github/workflows/ci.yml` 的 GitHub push/PR 最终绿灯仍是发布前外部条件；本地 build、Unit、Integration、Docs 与 API compare 结果不能替代 CI。
- 生产仓库入口仍不在本仓库。2C4G 部署必须在进入 NPOI DOM 前共享 `SemaphoreSlim(1, 1)`，保证最大实际并行度为 1；实际入口未确认或未落实该门禁时仍为 No-Go。
- 本轮 Reviewer 仅更新 `review.md`，未修改测试、生产代码、计划、执行报告、报告证据、baseline、依赖或 CI，未执行 git add/commit/push/PR/tag/publish。
