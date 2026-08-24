# TASK-BING-OFFICES-20260821-MAPVAL-V2 最终报告

## 状态

- Task ID：`TASK-BING-OFFICES-20260821-MAPVAL-V2`
- 最终状态：`PARTIAL`
- Compatibility Mode：`MIGRATION_CURRENT_MAJOR`

## 已实现

- 双模型 Import/Export Mapping Profile、方向隔离 Builder、Profile Registry、显式注册和程序集扫描。
- v2 JSON/XML normalized document，v1 平铺配置兼容适配，JSON/XML 输入限额、未知字段拒绝、XML DTD/外部实体防护和流所有权保持。
- Required、Regex、Date、MaxValue、Range、Unique、MaxLength 新 Attribute；Regex timeout、独立错误码和 Unique committed/pending journal 已进入 CSV/NPOI 主链。
- CSV/NPOI 方向化 Mapping Document 接入，固定列/动态列、Alias、Named Converter/Validator 和资源上限覆盖已有主链。
- 生产 IVT 已删除，仅保留测试友元；Public API allowlist/hash、provider-neutral consumer 和三包本地 pack consumer 已验证。
- 文档和可编译 Docs Consumer 已完成；文档示例已修正为真实 `HasHeader` API。
- 文件 JSON/XML 入口改为受限流读取，避免 `ReadAllText` 绕过输入大小上限。

## 验证

- Release solution build：通过，0 错误，180 个警告。
- Unit：net6 `156/156`，net8 `156/156`。
- Integration：net6 `10/10`，net8 `10/10`。
- Docs Consumer：net8 `3/3`。
- Loader 安全专项：`8/8`。
- 三包 pack 和本地隔离 consumer：通过，运行输出 `pack-consumer-ok`。
- Benchmark ShortRun：1K Import `15.90 ms / 16.36 MB`，Export `15.68 ms / 8.26 MB`；存在 GC 回收，未宣称零 GC。

## 未完成与风险

- 计划要求的完整 immutable `WorkbookPlan/SheetPlan/ColumnPlan` provider-neutral compiler 尚未形成；当前仍以 `ExcelTypeMap` 和 NPOI 列计划适配为主。
- 尚无正式 diagnostic/round-trip loader API、model alias registry、完整精确 JSONPath/XML path 错误模型。
- Unique FirstRowNumber 错误元数据和公开 comparer 选项尚未闭环。
- Provider SPI 仍不够窄；生产 IVT 虽已清理，但内部类型映射调用链尚未完全迁移到只读 Plan view。
- 性能矩阵仅覆盖现有 1K Excel Import/Export；LOH 和 Peak Working Set 未测量；外部 Office/LibreOffice 重开未执行。
- locked restore 的历史 `NU1004` 仍存在，执行使用 `--force-evaluate`，未改写受跟踪 lock 文件。

## 迁移说明

当前 major 保留 `ExcelMappingProfile<T>`、旧 Attribute、`ExcelMappingConfiguration`、Loader 旧重载和 `AddNpoi(): void`。新代码应使用双模型 `ExcelMappingProfile<TImport,TExport>`、v2 Document、新 Attribute 和方向化配置。旧 API 仅作为 Obsolete 兼容入口，不在本 Task 删除。

## 变更范围

修改集中在 `src/Bing.Offices.Abstractions`、`src/Bing.Offices.Core`、`src/Bing.Offices.Npoi`、Unit/Docs Tests、`docs/excel` 和本 Task 证据目录；用户原有 13 个 `Metadata/Excels` 删除项已保留。

明确声明：未执行 Commit、Push、PR、Tag 或 NuGet 发布。