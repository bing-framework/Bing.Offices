<!-- AI_REVIEW_STATUS: PASS_WITH_ISSUES -->
AI_TASK_ID: BO-FUNCTIONAL-GAPS-20261024-001
AI_REVIEWED_AT: 2026-10-24T23:06:31+08:00

# 修复与优化复核（Round 2）

## 结论

P2 回归 FIX-001 已修复并通过本轮复核。本轮覆盖原 5 项修复、P2 更正、相关 CSV/NPOI 调用路径和工作流；未发现新的 MUST_FIX / SHOULD_FIX。存在 4 项可选优化，当前结论为 PASS_WITH_ISSUES。生产发布仍依赖远端 CI 和目标环境证明。

本轮依据用户要求、execution.md、当前差异、真实代码和验证产物。没有 plan.md，未进行正式基于计划的验收。上一轮审查完整保存在 `review-round-1.md`，本文件为最新复核结论；execution.md 和业务代码未改写。luna_worker 独立复核了 P2 代码、8 项用例、双 TFM TRX 和 CI 重复调用。

```text
IMPLEMENTATION_STATUS: PASS
TEST_STATUS: PASS（复核已有证据，未重新运行测试）
PLAN_ACCEPTANCE_STATUS: NOT_APPLICABLE
PERFORMANCE_EVIDENCE_STATUS: NOT_MEASURED（优化收益未做耗时测试）
EXTERNAL_GATE_STATUS: BLOCKED_EXTERNAL
RELEASE_STATUS: BLOCKED
OPEN_ACTIONABLE（必须修复）: 0
OPTIONAL（本轮仅提出）: 4
BLOCKED_EXTERNAL: 2
```

## 修复验收

| 修复项 | 状态 | 本轮核对 |
| --- | --- | --- |
| GAP-001：重复表头绑定 | CLOSED | CSV 和 NPOI 分开检查固定属性与动态 Key，仍拒绝真正重复目标。 |
| FIX-001：Name/name 误拒绝 | CLOSED | NpoiImportSheetExecutor.cs:185 使用 StringComparer.Ordinal；表头查找仍使用请求 HeaderComparison。新测试覆盖 XLS/XLSX、同步/异步、OrdinalIgnoreCase/Ordinal 的 8 种组合；忽略大小写模式用 lower-case 标题验证匹配策略仍生效。 |
| GAP-002：无表头 CSV 动态列 | CLOSED | 位置绑定与导出均按属性顺序展开，并以 Order、Key 排序动态列；缺失动态定义明确报错。 |
| GAP-003：CSV 字段限制 | CLOSED | 同步/异步解析器均使用原始转义长度上限，解码后逐字段检查，覆盖表头和未映射字段；资源错误转换路径仍存在。 |
| GAP-004：覆盖率门禁 | CLOSED | collector、真实 XML 产物及缺失/空报告失败检查保留。 |
| GAP-005：文档示例一致性 | CLOSED | 生成 Markdown 后运行 docs-check，原正、负向证据仍适用。 |

证据来自 `TestResults/p2-fix/unit/` 两份独立 TRX：net6.0 / net8.0 各 795 passed、0 failed、0 notExecuted，共 1590 项；`Import_CaseDistinctFixedProperties_ShouldBindBothValues` 在两个 TFM 各 8 项全部通过。修复前 TRX 中该测试为 4 failed / 4 passed，已形成前后对照。当前 NPOI 表头绑定类为 32 项，保留原来的 24 项成功及失败场景。

本轮复用了最近的双 TFM 单元测试和未改变范围的文档/覆盖率证据，未重复跑全解决方案、Benchmark 或生产资源矩阵。静态差异检查通过；所查 CSV 与 NPOI 表头路径未发现新增 TODO/NotImplemented、Task.Run、Result 或 Wait 伪异步实现。不将这一局部复核解释为整个仓库不存在任何缺陷。

## 可选优化

以下均为 LOW / OPTIONAL，状态 DEFERRED：本轮用户要求检查，尚未实施；不作为修复或发布阻塞项，也不触发自动 Fix。建议后续独立任务为 `BO-IMPORT-CI-OPTIMIZATION`，先做可直接减少重复工作的两项，再决定内部结构优化。

### OPT-001：合并普通 net8 测试与覆盖率采集

- 证据：`.github/workflows/ci.yml:40` 与 `49`，`.github/workflows/docs.yml:48` 与 `61`，均对相同的 net8 单元测试项目执行两遍，第二遍用于覆盖率。
- 影响：按当前用例数，每份工作流多执行 795 项 Unit 用例；未测量 CI 墙钟耗时，不能据此承诺提速比例。
- 建议：将 `--collect:"XPlat Code Coverage"` 合并进原 net8 测试调用，保留 TRX、非空 Cobertura 门禁和 artifact 归档，再删除重复调用。
- 实施约束：Docs 的 TRX collector 需要 `TestResults` 顶层 TRX；覆盖率文件可位于 GUID 子目录，必须同步调整递归检查路径及相关 README，不能只删除第二次测试命令。
- 验证：同一轮同时生成 TRX 与非空可解析 Cobertura；缺失/空报告仍失败；文档能采集测试与覆盖率数据。

### OPT-002：已完成 solution 构建后，测试使用 --no-build

- 证据：`.github/workflows/ci.yml:31` 已构建完整 Release solution；第 39～43 行的 5 条测试命令未设置 --no-build。测试项目均在 solution 的 Release Build.0 列表中。覆盖率调用已经设置此选项。
- 影响：多出 5 次测试项目/TFM 的 build target 调用；增量构建不等于完整重编译，收益取决于实际 CI 环境。
- 建议：为这 5 条调用添加 --no-build，保留前置 build 的失败门禁与所有测试。
- 验证：从干净检出开始，前置 build 必须产出所需双 TFM 和 Docs 测试 DLL，后续测试仍完整执行。
- 边界：Docs 的两个 job 位于独立 runner，各自 build 不据此判为冗余；合并会涉及 job 隔离和状态检查变化，不纳入这一简单优化。

### OPT-003：统一 CSV 解析配置的构造

- 证据：`CsvPipelineSupport.cs:48` 与 `81` 重复构造 CsvConfiguration，包括分隔符、引用字符、RFC4180、MaxFieldSize 和坏数据回调。
- 影响：当前两处一致，没有功能错误；未来修改需要维护两份，存在同步/异步配置偏移的维护风险。
- 建议：提取私有配置工厂，每次读取创建独立配置，不引入共享可变配置或新缓存。真实读操作仍分别使用 Read / ReadAsync。
- 验证：复用同步/异步长度限制、引号/换行、取消和异步流用例；不声称这一重构能显著提高吞吐。

### OPT-004：宽表场景评估表头查找索引

- 证据：`CsvHeaderBinder.cs:55,62,82` 每个表头扫描属性或动态定义，再扫描绑定列寻找缺失属性；NPOI 的 `CreateColumns`、`FindProperty` 和 `FindDynamicDefinition` 也使用逐项查找。
- 影响：忽略别名展开长度时，查找约为 O(H×(P+D))，缺失检查另有 O(P×H)；H/P/D 分别为表头、固定属性和动态定义数。该过程按输入或 Sheet 绑定一次，不是每行执行，不能据此断言普通数据量存在瓶颈。
- 建议：如果实际业务使用数百/数千列，先用代表性宽表测量绑定阶段，再决定是否建立本次绑定内的索引。没有测量依据前不新增全局计划缓存。
- 实施约束：必须保留现有标题/别名/属性名的匹配优先级、表头比较策略、Name/name 属性身份以及冲突处理；不能直接 ToDictionary 覆盖冲突项。
- 验证：对照完整绑定结果与错误结果，测量绑定时间和分配；保留重复表头、动态别名及 Name/name 回归。

## 外部项与停止条件

- EXT-001 / BLOCKED_EXTERNAL：当前候选的远端 GitHub Actions 绿色结果尚需外部运行。
- EXT-002 / BLOCKED_EXTERNAL：目标 2C4G/semaphore=1 部署与资源预算证明尚需目标环境提供。
- 本轮必修问题 0；本轮只形成优化建议，不修改实现、不执行自动修复循环、不提交或推送 Git。
- 下一步：STOP。若开展优化，优先 OPT-001、OPT-002；OPT-004 应先取得实际宽表数据的测量结果。
