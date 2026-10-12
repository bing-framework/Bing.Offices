<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001
AI_EXECUTION_FINISHED_AT: 2026-10-03T14:10:29Z

AI_REVIEW_FIX_STARTED_AT: 2026-10-03T13:36:45.778Z
AI_EXECUTION_STARTED_AT: 2026-09-30T16:01:14.030Z

# 执行记录

## 实施结果

本轮完成实体布局分批导入。NPOI 和 ClosedXML 复用既有完整实体导入器，完成 DOM 导入后按布局列表区域投影批次；没有改变原有实体导入、模板、校验、关系、动态列、Footer、分页/小计或流所有权行为。

## 变更范围

- Abstractions 新增 `IExcelEntityBatchImporter`、`ExcelEntityBatchImportOptions`、`ExcelEntityImportBatch<TEntity>`、`ExcelEntityBatchRegionSummary` 和 `ExcelEntityBatchImportSummary<TEntity>`。
- Provider SPI 支持普通布局和模板布局的同步/异步入口；NPOI 支持 XLS/HSSF 与 XLSX/XSSF，ClosedXML 支持 XLSX。
- 新增共享 `ExcelEntityBatchProjection`，按 `SourceRows` 与实体集合配对，跨区域维护全局批次号；固定单元格、关系和其他实体级错误进入 `RegionIndex = -1` 的诊断批次。
- NPOI/ClosedXML 完成 DI 注册并声明 `ExcelProviderFeatures.EntityLayoutBatchImport`；MiniExcel、SpreadCheetah、ExcelDataReader 未伪造该能力。
- 更新 Provider 能力矩阵、Excel README、API 审批记录和生产符号到测试方法追溯。
- 未加入 Entity Failure Workbook 选项；实体完整结果的结构化错误和诊断批次是本轮错误出口。

## 测试证据

- NPOI `EntityLayoutBatchImportTest`：net6.0 `6/6`，net8.0 `6/6`；覆盖 XLS/HSSF、XLSX/XSSF、双列表区域、批次边界、固定字段错误诊断、模板异步回调、预取消和流所有权。
- ClosedXML `EntityLayoutBatchImportTest`：net6.0 `4/4`，net8.0 `4/4`；覆盖 XLSX 双列表区域、固定字段错误诊断、模板异步回调、预取消和流所有权。
- NPOI 与 ClosedXML Release 项目构建：双 TFM 均成功，0 错误。
- NPOI Provider 全量 `Category!=Large`：net6.0 `806/806`，net8.0 `806/806`。
- ClosedXML Provider 全量 `Category!=Large`：net6.0 `255/255`，net8.0 `255/255`。
- 解决方案 Release 构建成功，0 错误；保留 4 个既有 Microsoft.Bcl.Memory net6.0 兼容性警告。
- 解决方案常规 `Category!=Large` 两次运行都未一次通过：首次 NPOI net8.0 资源探针清理共享 `zip.xlsx` 时发生文件锁错误；第二次 NPOI net6.0 异步暂存清理测试的共享临时目录集合断言失败。两项分别独立重跑通过，NPOI net8.0 完整项目通过；失败是否由其他进程并发造成尚无确定证据。
- 公共 API 合同：net6.0 `11/11`，net8.0 `11/11`；Abstractions 双 TFM 仅新增 44 个成员，删除 0 个，已核对并更新 baseline。
- 既有 Microsoft.Bcl.Memory 对 net6.0 的兼容性警告继续存在，未由本轮引入。

## API 与兼容性

- `api-approval.md` 记录本轮新增成员及迁移边界。
- `artifacts/api-compare/api-diff.json` 保存双 TFM 成员差异；仅 `Bing.Offices.Abstractions` 有新增，其他程序集无差异。
- `build/api-snapshot-baseline.json` 已由当前候选快照更新，保留既有候选身份字段和批准元数据格式。
- 既有 `IExcelEntityImporter`、实体布局构造函数和完整导入入口保持不变。

## 待验证与限制

- Entity 批次是完整 DOM 导入后的投影，不提供前向读取或低内存承诺。
- Entity Failure Workbook 仍属于后续能力；当前通过 `ExcelEntityBatchImportSummary<TEntity>.Result` 与诊断批次返回错误。
- 本轮没有修改版本号、提交、推送、发布或外部 `Utopa.Erp.IE`。

## 收口检查

- [x] 运行受影响 Provider 全量 `Category!=Large` 测试。
- [x] 运行解决方案 Release 构建与常规 `Category!=Large` 测试；如实记录两次非全通过结果和独立复核结果。
- [x] 执行 `git diff --check` 及改动文件 UTF-8 BOM/EOL 检查。
- [x] 将第一行状态更新为 `COMPLETED` 并运行任务状态收口。

## Review 修复记录

### Round 1

- Review 状态：NEEDS_FIX；Fix Scope：recommended。
- Review 文件：`ai_docs/tasks/BING-OFFICES-ENTITY-LAYOUT-BATCH-20260930-001/review.md`，保留原始结论供独立复核。

#### FIX-001：Entity 批次组合与失败边界

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- NPOI 与 ClosedXML 的 `EntityLayoutBatchImportTest` 增加真实工作簿用例，断言动态列、关系、空列表与无列表区域、失败行错误归属、分页小计及 Footer 后的原始行号、回调异常/取消和 getter-only 集合内容。
- 直接测试：NPOI net6.0/net8.0 各 18/18；ClosedXML net6.0/net8.0 各 11/11。NPOI 用例包含 XLS/HSSF 与 XLSX/XSSF。
- 新用例揭示 ClosedXML 对 getter-only 可写集合未填充的问题；`ClosedXmlEntityLayoutExecutor` 现与 NPOI 一样填充现有集合，并在不可写时返回配置异常。

#### FIX-002：Entity 批次包消费者

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- net6.0、net8.0 Consumer 均从临时版本 `2.0.0-entitybatch-review2` 的 NuGet 包解析 NPOI 和 ClosedXML `IExcelEntityBatchImporter`，导入真实双 Sheet 工作簿，核对固定字段诊断批次、成功明细批次及摘要；两者均运行通过。
- 最终包 SHA-256：NPOI `9172461B16A0A96F30BC3E619E465CCD2C0969FB08BE444EF7B8A70C131E20B0`；ClosedXML `3BCE04EC01E166093A3869FBDD5672C3C21C8337939E7BDF3C3D6698CF0753F1`。
- 双 TFM 公共 API 合同包含在本轮解决方案常规测试中，未修改公共 API 或批准快照。

#### FIX-003：文档与最终验证证据

- 严重程度：MEDIUM；处理要求：SHOULD_FIX；执行状态：COMPLETED。
- 更新 Provider 能力说明与生产符号到测试方法映射，明确空区域、错误行、动态列、关系、分页小计及 Footer 等对应的职责测试。
- 旧执行记录保留了此前两次未一次通过的历史结果，且已删除未经证实的并发残留归因。本轮使用项目及 TFM 串行设置，Release 解决方案构建通过（0 错误、4 条既有 net6.0 依赖警告）；`Category!=Large` 一次运行通过，共 3775 个测试、0 失败。该结果是本轮新鲜门禁，不将此前独立重跑冒充为旧门禁通过。
- `git diff --check` 通过；既有 ProfileFixtures XML 的 CRLF 提示不属于本轮编辑。

### Round 1 汇总

- MUST_FIX：0；SHOULD_FIX：3 项完成；PARTIAL/BLOCKED/FAILED：0。
- 回归验证：NPOI 与 ClosedXML 双 TFM 职责测试、解决方案 Release 构建、串行常规测试及双 TFM 包消费者均通过。
- 保留边界：Entity 批次仍是完整 DOM 导入后的投影，不提供低内存前向读取；分组分页组合、重复页命名锚点与公共 PivotTable 仍属后续独立设计。商业环境公式重算与渲染验收仍需实际许可证和部署环境。
- 下一步：独立 Review 本轮修复；不自动提交、推送或发布。
