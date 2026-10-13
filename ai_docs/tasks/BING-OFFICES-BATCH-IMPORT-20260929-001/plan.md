# ExcelDataReader 多 Sheet 分批导入扩展

## 目标

在不改变现有 `IExcelBatchImporter` 行为和签名的前提下，为 ExcelDataReader 增加 Workbook 级多 Sheet 分批导入契约。调用方可以一次暂存输入工作簿，按源文件物理顺序收到异构 Sheet 的同步或异步批次，并通过批次中的 Sheet 名称和实体类型保持错误和持久化路由。

## 范围

1. 在 Abstractions 中新增 `IExcelWorkbookBatchImporter`、`ExcelWorkbookImportBatch`、工作簿批次摘要及 Sheet 摘要类型；为 `ExcelImportBatch<TItem>` 增加来源 Sheet 只读属性，并保留旧构造函数。
2. 在 ExcelDataReader 中实现同步、异步多 Sheet 批次读取。每个 Sheet 使用既有固定列 Mapping、Converter、配置校验、资源限制、取消、回调异常和临时文件生命周期契约；工作簿输入只暂存一次，按请求独立重置前向读取器。
3. 新增 `ExcelProviderFeatures.MultiSheetBatchImport`，仅 ExcelDataReader 声明；DI 注册新接口。其他 Provider 不伪造能力。
4. 通过职责级测试覆盖 XLSX 多 Sheet、批次顺序、来源 Sheet、同步/异步回调、取消、错误与资源边界、输入流所有权、旧单 Sheet API 回归和 DI 注册。
5. 更新 Provider 文档、能力矩阵、API 审批、双 TFM snapshot、生产符号到测试方法映射和执行记录。

## 明确边界

本轮工作簿批次请求支持 Workbook 请求中不同 Sheet 的实体类型，回调通过 `ItemType` 和 `Items` 进行安全分派。动态列、跨 Sheet 关系、失败工作簿、Workbook 原生校验、跨行唯一性继续按现有分批路径在预检阶段拒绝，后续单独设计这些能力的组合。

## 验收

- 旧 `IExcelBatchImporter` 和 `ExcelImportBatch<TItem>` 四参数构造调用继续编译并保持结果。
- 新接口在 XLSX 的同步、异步多 Sheet 场景交付完整且顺序稳定，批次错误带正确 Sheet 名称。
- 预取消、回调取消/异常、资源限制后目标输入流仍由调用方拥有，暂存文件清理。
- net6.0/net8.0 Release 受影响项目、Provider 合同、API snapshot 和常规 `Category!=Large` 检查通过；不更新基线掩盖未批准差异。
