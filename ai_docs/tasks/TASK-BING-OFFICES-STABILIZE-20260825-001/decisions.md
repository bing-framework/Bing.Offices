# 决策记录

## D-001 ValidationMode 分流

配置规则只在 `ConfiguredRules` 和 `ConfiguredAndWorkbook` 启用；Workbook 原生规则只在 `WorkbookRules` 和 `ConfiguredAndWorkbook` 启用。Workbook 校验保持先执行；`ValidateMode.Continue` 下继续配置校验以收集完整错误，但 Workbook 不通过的实体不会提交。

## D-002 失败工作簿 Sheet identity

失败产物使用导入阶段已解析的实际 Sheet 名称到请求的映射。这样 `ByIndex` 不依赖构建时的 `#N` 名称，也不会在多 Sheet 输出时通过用户 selector 二次猜测。

## D-003 模板单元格覆盖策略

默认 `PreserveTemplate`：写入导出值时保留模板样式和批注，已有公式被导出值替换。`ReplaceTemplate` 是显式选择，写入前清除目标单元格样式和批注。当前策略不承诺公式重写或行块复制。

## D-004 Mapping 合并边界

后续实现以 `ExcelMappingPlanFactory` 作为最终合并点。Builder 和 DocumentFactory 只保存/克隆层级快照，不提前把 Document 与 Request 合并；Provider 仅消费最终计划。

## D-005 方向缺失与显式 fallback

`ExcelMappingDocument.Import`/`Export` 的 null 表示方向未提供。Plan Factory 在目标方向、Request configuration 均缺失且 `UseConventionFallback=false` 时抛出明确异常；无 Document 的传统 CSV/NPOI 路径显式创建 `UseConventionFallback=true`，避免隐式静默回退。

## D-006 Profile 只读解析与稳定名称

新增 `IMappingProfileResolver` 作为计划编译和业务读取边界，`IMappingProfileRegistry` 继续承担启动期 `Register`。提供 `AddMappingProfile<T>(string profileName)` 稳定名称注册重载；未提供名称时保留 `FullName` 兼容行为。Profile 扫描仍在 Npoi，Core 迁移留待后续阶段。

## D-007 Benchmark 证据边界

Mapping 基准只测计划、解析、校验和注册路径，不主动分配 90KB payload 伪装产品 LOH 证据。LOH/working set 压力场景保留在独立 ResourceProbe；计划构建基准分别报告 cold、cache hit、cache miss。
