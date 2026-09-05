# 执行进度

| Task | 状态 | Scope | Evidence | Risk | Remaining | Updated |
| --- | --- | --- | --- | --- | --- | --- |
| RC0-01 | VERIFIED | 环境、Git、项目和规则基线 | `baseline.md` | 共享工作区含前序任务差异 | 无 | 2026-09-04 |
| RC0-02 | VERIFIED | Restore、Build、Unit、Integration、Docs、API、Pack | `unit-test-report.md`、`integration-test-report.md`、`docs-test-report.md`、`api-diff.md` | formal hash 未批准 | API approval | 2026-09-04 |
| RC0-03 | VERIFIED | API、异常、日期、NPOI 扩展、弃用矩阵 | `api-diff.md`、`deprecated-removal.md` | public surface 仍需治理 | 逐符号审批 | 2026-09-04 |
| RC1-01 | VERIFIED | 专属异常合同和单次翻译 | Unit/Integration reports | public dispatcher 长期治理待定 | API approval | 2026-09-04 |
| RC1-02 | VERIFIED | Excel/CSV 确定性日期解析 | Unit/Integration reports、package consumer | 完整真实文件日期矩阵不足 | 补充矩阵 | 2026-09-04 |
| RC1-03 | VERIFIED | API Contract 单一事实来源 | `PublicApiContractTest`、`api-diff.md` | formal hash 仍不匹配 | API approval | 2026-09-04 |
| RC1-04 | VERIFIED | XLSX ZIP 预检和 XLS/OLE 策略 | ZIP Unit、Excel ResourceProbe、`resource-report.md` | XML 深度、XLS/OLE、DOM 硬上限不足 | 资源预算审批 | 2026-09-04 |
| RC1-05 | VERIFIED | Failure Workbook/Data Validation | Unit/Integration reports | 双 DOM 资源矩阵不足 | 补充资源矩阵 | 2026-09-04 |
| RC2-01 | VERIFIED | NPOI public extensions | API contract、package consumer | Provider API 需批准 | API approval | 2026-09-04 |
| RC2-02 | VERIFIED | Mapping 死字段和入口 | `deprecated-removal.md`、Mapping tests | remaining public execution detail | 逐符号审批 | 2026-09-04 |
| RC2-03 | VERIFIED | 含糊命名迁移 | `deprecated-removal.md`、Unit、consumer | breaking migration 需批准 | API approval | 2026-09-04 |
| RC2-04 | PARTIAL | DataTable/旧 helper 删除 | `deprecated-removal.md` | 显式 DataTable API 外部价值未裁决 | compatibility decision | 2026-09-04 |
| RC3-01 | PARTIAL | 异常目录边界 | 源码审计、API diff | execution detail 仍 public | 逐符号治理 | 2026-09-04 |
| RC3-02 | PARTIAL | Failure Workbook 拆分 | Failure Workbook tests | 计划中的职责拆分未完成 | 后续重构任务 | 2026-09-04 |
| RC3-03 | PARTIAL | CSV writer 生命周期 | CSV tests、benchmark report | 当前任务未完成 A/B 矩阵 | 专项 benchmark | 2026-09-04 |
| RC3-04 | PARTIAL | Mapping compiler/index | Mapping resource probe、benchmark | 完整 10K rule A/B 未完成 | 专项 benchmark | 2026-09-04 |
| RC3-05 | PARTIAL | Public/Internal extensions 目录 | API contract、consumer | 目录整理未执行 | 后续重构任务 | 2026-09-04 |
| RC4-01 | VERIFIED | Unit P0/P1 矩阵 | `unit-test-report.md`、3 TFM TRX | formal API hash 失败 | API approval | 2026-09-04 |
| RC4-02 | VERIFIED | Integration | `integration-test-report.md`、net6/net8 TRX | Windows-only、真实日期矩阵不足 | 扩展平台/日期矩阵 | 2026-09-04 |
| RC4-03 | VERIFIED | Docs/XML docs | `docs-test-report.md`、11/11 TRX | 最终资源/API 文档待批准 | 文档收口 | 2026-09-04 |
| RC4-04 | VERIFIED | Package-only consumer | `package-consumer-report.md`、net6/net8 | feed/clean-clone 未验证 | 发布环境复现 | 2026-09-04 |
| RC5-01 | PARTIAL | 基准有效性 | `benchmark-report.md` | 历史 100k 3 异常未解释 | 重新定位基准 | 2026-09-04 |
| RC5-02 | PARTIAL | Benchmark 矩阵 | `benchmark-report.md`、BDN results | 仅 MappingValidation 10/10 | 完整矩阵/预算 | 2026-09-04 |
| RC5-03 | PARTIAL | 资源探针 | `resource-report.md`、Excel 12 + mapping 16 | Failure Workbook/DOM/取消矩阵不足 | 扩展 probe | 2026-09-04 |
| RC5-04 | PARTIAL | 性能资源报告 | `benchmark-report.md`、`resource-report.md` | budgets 未批准 | 维护者预算 | 2026-09-04 |
| RC6-01 | PARTIAL | 文档和迁移 | `deprecated-removal.md`、Docs report | 任务专属资源/API 文档仍待批准 | 最终文档审阅 | 2026-09-04 |
| RC6-02 | BLOCKED | API diff 和批准 | `api-diff.md` | breaking diff 无本任务批准 | 维护者批准并重跑 | 2026-09-04 |
| RC6-03 | VERIFIED | 独立 Review | `review.md`、Round 1/2 记录 | Review 修复后仍需最终 sign-off | 维护者 sign-off | 2026-09-04 |
| RC6-04 | BLOCKED | 最终发布门禁 | `final-report.md` | API、性能、资源和交付条件未满足 | 解除全部阻断 | 2026-09-04 |
