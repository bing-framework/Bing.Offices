<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-LEGACY-MIGRATION-DIAGNOSTICS-20260928-001
AI_EXECUTION_FINISHED_AT: 2026-09-28T13:10:54.000Z
AI_EXECUTION_STARTED_AT: 2026-09-28T13:04:11.878Z
AI_EXECUTION_RESULT: PASS

# 旧版 Excel API 迁移诊断：执行报告

## 执行结论

新增 `BOM001` 和 `BOM002` 两条只读迁移警告，仅识别已绑定到明确旧版完整类型的代码。不自动转换旧选项、动态列或业务自定义坐标属性。

## 任务信息与影响分析

- 计划：[plan.md](plan.md)。外部 `Utopa.Erp.IE` 只作只读样本，未修改或复制其业务源码。
- ChangedProjects：可选分析器、分析器职责测试、NuGet 包消费者和 CI；Office 运行时 Provider、公共请求及文件格式不变。
- ChangedPublicContracts：分析器新增 `LegacyApiMigrationAnalyzer` 和规则 ID `BOM001`/`BOM002`；八个运行时生产程序集无公共 API 变化，API snapshot 无需更新。
- 风险：旧类型若在升级后无法绑定，诊断不会猜测；与 2.x 共存的 `ColumnNameAttribute` / `DynamicColumnAttribute` 不告警。

## 计划执行情况

1. `BOM001` 对 `Bing.Offices.Exports.IExcelExportService`、`ExportOptions<T>` 的类型引用发出 Warning。
2. `BOM002` 对 `Bing.Offices.Attributes.HasDynamicColumnAttribute` 的使用发出 Warning。
3. 单元测试覆盖业务样本常见的 `using` 用法、特性、同名业务类型与保留特性；本地 NuGet 包消费者出现预期警告。
4. 迁移文档说明旧服务、多个动态字典和自定义 `ImportLocationAttribute` 的人工转换边界；CI 增加真实包消费检查。

## 部分或未完成事项

本阶段无未完成实施项。自定义 `ImportLocationAttribute` 不属于通用旧 API，缺少可保证正确性的自动转换规则，保留为人工迁移。

## 修改文件

分析器 `LegacyApiMigrationAnalyzer.cs`、规则清单、职责测试 `LegacyApiMigrationAnalyzerTest.cs`、包消费者 `LegacyMigration.cs`、CI 工作流、分析器及 NuGet 迁移文档、主任务索引与追溯。

## 测试与构建

- 分析器全部职责测试：net6.0 20/20、net8.0 20/20。
- 本地打包 `Bing.Offices.Analyzers.2.0.0.nupkg` 后，用隔离缓存还原消费者；条件编译样例得到两条 `BOM001`、一条 `BOM002`，0 错误。
- Docs Markdown 代码围栏测试 net8.0 1/1。构建保留既有 NuGet 源不可达和 CodePages 7.0.0 回退到 8.0.0 的警告。
- 未运行无关 Provider 全量、API snapshot 或性能矩阵；CI 新步骤尚未在 GitHub Actions 执行。

## TODO 状态

- Completed：两个旧 API 诊断、双 TFM 职责测试、真实包消费者、CI 和迁移文档。Implementation 4/4。
- Open Actionable：无。Next Action：STOP。
- Blocked Approval：无。
- Blocked External：GitHub Actions 尚未执行；本地同构包消费者已通过。
- Not Applicable：运行时 API 差异、Provider 资源矩阵。
- Accepted Limitations：诊断要求旧类型仍可绑定；业务自定义坐标特性须人工审查。
- Verified Boundaries：无新增资源边界。
- Deferred：Source Generator 仅在有反射成本证据后评估；不增加无依据的旧 API 自动 Code Fix。

未自动 git add、commit、push、创建 PR、发布或修改版本号。
