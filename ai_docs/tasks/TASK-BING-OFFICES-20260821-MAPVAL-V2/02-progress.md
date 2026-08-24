# TASK-BING-OFFICES-20260821-MAPVAL-V2 执行进度

## 当前状态

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 主阶段：Phase 6，`PARTIAL`
- 当前任务：MAPVAL-602，`PARTIAL`

## 已执行

1. 读取项目规则、Skill、计划和现有 Excel 设计/测试/CI。
2. 注册 `plan-execution` 运行状态，source 为 `copilot`。
3. 采集 Git 分支、HEAD 和工作区差异。
4. 执行 locked restore，确认现有 lock 文件与项目依赖表达式不一致。
5. 使用 `--force-evaluate` 完成还原；Release Build、Unit、Integration、Docs Consumer 基线全部通过。

## 发现

- 当前版本为 `1.0.0`，采用 `MIGRATION_CURRENT_MAJOR`。
- 当前已有 Workbook Request、动态列、读取范围、空白策略、图片/失败输出/部分 Workbook Validation。
- Mapping Profile、统一 Core/Npoi/CSV Plan、v2 JSON/XML、六类新校验、Profile Registry 和生产 IVT 清理仍未完成。
- 用户已有 Metadata/Excels 删除项不属于本 Task，必须保留。
- 现有 `ExcelMappingProfile<T>` 只保存平铺 configuration；现有 Import/Export Request 已分别持有该旧 Profile 类型。
- 现有 DI 入口位于 Npoi，`AddNpoi()` 返回 `void`，必须保持该签名。

## 本轮增量

- 已完成方向化 Mapping Profile、Registry 显式/扫描注册、CSV/NPOI Unique pending journal 和 v2 校验 Attribute 接入。
- 已新增 `ExcelMappingDocument`，旧 JSON/XML 平铺格式经 v1 adapter 归一化；新增 Loader document API 保留原有 configuration API。
- 已加入 JSON 深度/大小限制、严格语法选项、XML DTD/外部实体禁用和 XML 文档字符上限。
- 已更新测试替身及 Public API allowlist/hash；未修改 `plan.md`，未执行 commit/push/PR。

## 本轮验证

- Full Unit net8：148/148 通过（最近完整结果）。
- CSV：22/22 通过；Excel P0/Workbook/Stream：113/113 通过；Registry：3/3 通过。
- Loader/API targeted 在新增接口替身补齐后已编译；Public API hash 已更新为实际快照，待严格重跑确认。

## 本轮收口

- 文档示例已修正为真实 `HasHeader` API，Docs Consumer 扩展至 3/3。
- 文件 JSON/XML 入口改为受限流读取，并新增 oversized file 回归用例。
- 三包本地 pack、隔离 package consumer restore/build/run 已通过，输出 `pack-consumer-ok`。
- BenchmarkDotNet ShortRun 已完成：1K 行 Import 平均 15.90 ms、16.36 MB；Export 平均 15.68 ms、8.26 MB；存在 Gen0/Gen1/Gen2 回收，未宣称零 GC。
- 最新 Unit net6/net8 均为 156/156；Integration net6/net8 均为 10/10；Docs Consumer 为 3/3。

## 明确遗留

- 尚未形成计划要求的统一 immutable `WorkbookPlan/SheetPlan/ColumnPlan` provider-neutral compiler；当前主链仍以 `ExcelTypeMap` 和 NPOI `ExcelColumnPlan` 适配为主。
- 尚无正式 diagnostic/round-trip loader API、model alias registry 和完整 JSON/XML 精确路径模型。
- `UniqueTracker` 已使用 committed/pending journal 和资源上限，但 FirstRowNumber 错误元数据及公开 comparer 选项尚未闭环。
- Provider SPI 仍不够窄，生产 IVT 已移除但 Core/Npoi 仍通过现有类型映射调用链协作。
- Benchmark 目前只有现有 1K Excel Import/Export；10K/100K、Unique 多列、Plan Build、租户缓存、JSON/XML 和注册扫描矩阵尚未新增。