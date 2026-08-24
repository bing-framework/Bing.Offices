# TASK-BING-OFFICES-20260821-MAPVAL-V2 决策记录

## ADR-001 当前 major 采用兼容迁移

- 状态：`DONE`
- 决策：采用 `MIGRATION_CURRENT_MAJOR`。
- 证据：`version.props` 为 `1.0.0`；三个 PackageId 已存在；用户要求保留包身份、旧 API additive/Obsolete 迁移。
- 影响：不删除旧公开 API，不改变 `AddNpoi(): void`；旧入口委托统一 compiler。

## ADR-002 相邻计划不能覆盖当前源码事实

- 状态：`DONE`
- 决策：`ai_docs/tasks/bing-offices-excel-import-export-enhancement-v2/plan.md` 只作为历史线索；当前源码/测试/项目配置优先。
- 证据：当前 Npoi importer 已包含 ReadColumns、Whitespace、图片、Workbook Data Validation、Failure Writer 等实现。
- 影响：执行时对已存在能力做直接测试和接入，不重复实现或按历史“未实现”删除。

## ADR-003 用户已有删除项不回滚

- 状态：`DONE`
- 决策：保留 Task 启动前的 13 个 `Metadata/Excels` 删除，不对其恢复、清理或重构。
- 风险：完整解决方案构建可能受这些删除影响；若发生，记录为基线问题并判断是否与本 Task 相关。

## 待决策

- `packages.lock.json` 是否只需 `--force-evaluate` 更新，待 Phase 0 命令验证。
- Provider SPI 的最小字段集合，待 MAPVAL-001 消费者引用盘点后冻结。
- 是否存在真实 `AddBingOffices()` public API，待符号搜索确认。