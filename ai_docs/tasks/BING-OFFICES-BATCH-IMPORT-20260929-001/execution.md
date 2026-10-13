# 执行记录

## 状态

`COMPLETED`

本任务在不改变既有 `IExcelBatchImporter` 签名和单 Sheet 行为的前提下，完成了 ExcelDataReader 的 Workbook 多 Sheet 固定列分批导入扩展。

## 已实现

- Abstractions 新增 `IExcelWorkbookBatchImporter`、Workbook 批次、Sheet 摘要和 Workbook 摘要类型。
- `ExcelImportBatch<TItem>` 增加来源 `SheetName`，旧四参数构造函数继续保留；ExcelDataReader 的旧单 Sheet 批次现在填充实际 Sheet 名称。
- `ExcelSheetImportBuilder<TItem>.BatchSize` 支持 Workbook 请求逐 Sheet 配置批次大小。
- ExcelDataReader 同步和异步路径一次暂存输入，按物理 Sheet 顺序交付异构实体批次，维护全局批次号、首行号、共享行数/错误预算和输入流所有权。
- Workbook 资源超限批次保留实际物理 Sheet 名称和索引上下文；名称与索引同时指向同一物理 Sheet 时以 Plan 配置异常确定性拒绝。
- 预检生成的 Mapping Plan 在读取阶段复用，避免自定义计划工厂重复构建。
- 新增 `ExcelProviderFeatures.MultiSheetBatchImport`，仅 ExcelDataReader 声明并注册 `IExcelWorkbookBatchImporter`。
- 动态列、关系、Failure Workbook、Workbook 原生校验和跨行唯一性仍在预检阶段明确拒绝。
- 更新 Provider 文档、第三方 Provider 矩阵和生产符号到测试方法映射。

## Change Impact Analysis

- **L0 公共契约**：Abstractions 增加独立 Workbook 批次接口、结果模型和一个兼容属性/构造函数；API 审批记录见 `api-approval.md`。
- **L1 Provider**：仅 ExcelDataReader 新增 Workbook 批次实现和 DI 注册；旧单 Sheet 逻辑仅补充来源 Sheet 上下文。
- **L2 资源与异常**：复用现有暂存、取消、错误收集、回调异常包装和清理机制；Workbook 批次共享资源预算。
- **L3 文档与验证**：新增职责级边界测试、双 TFM 包消费者和 API 快照复核。

## 验证结果

- `dotnet build Bing.Offices.sln -c Release --no-restore -m:1`：通过，0 错误；仅有仓库已知的 net6.0 支持警告。
- `dotnet test Bing.Offices.sln -c Release --no-build --no-restore -m:1 --filter Category!=Large`：通过；所有已运行测试通过。
- ExcelDataReader 职责测试：net6.0 `44/44`，net8.0 `44/44`。
- Provider 合同测试：net6.0/net8.0 各 `298/298`。
- 公共 API 契约测试：net6.0/net8.0 各 `50/50`。
- Docs 测试：net8.0 `19/19`。
- API 快照已重新采集并与已审批基线比较，双 TFM 成员差异为零。
- 包消费者使用任务专用包源的临时验证版本 `2.0.1`：net6.0、net8.0 均运行成功；该版本只存在于任务产物，不修改仓库版本文件。
- `git diff --check` 和严格 UTF-8/BOM/EOL 检查在最终收口阶段执行；工作区既有 Profile Fixtures XML 的 CRLF 提示保留并单独记录。

## 未纳入本轮

Workbook 批次暂不支持动态列、跨 Sheet 关系、Entity Layout、Failure Workbook、Workbook 原生校验和跨行唯一性；这些能力按计划保留给后续独立扩展。未自动提交、推送或发布。
