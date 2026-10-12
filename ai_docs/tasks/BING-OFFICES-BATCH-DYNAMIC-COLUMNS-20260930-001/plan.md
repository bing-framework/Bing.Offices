# ExcelDataReader 分批动态列支持

## 目标

在不增加公共 API、不中断已交付固定列和 Workbook 原生校验行为的前提下，为 ExcelDataReader 的固定列分批接口补齐动态列读取能力。动态列复用既有 Mapping Plan、转换器、校验绑定和唯一性跟踪；每个实体独立持有动态字典，批次之间不共享可变字典状态。

## 范围

- `ImportBatches`、`ImportBatchesAsync` 支持带 `[DynamicColumn]` 字典属性的单 Sheet 请求。
- `ImportWorkbookBatches`、`ImportWorkbookBatchesAsync` 支持现有 Sheet Builder 的 `DynamicColumns(...)` 目标委托和映射计划动态定义。
- XLS、XLSX、XLSB 均沿用 ExcelDataReader 前向读取路径；XLSX Workbook 原生校验继续对动态单元格使用同一快照。
- 支持动态列标题、别名、数据类型、值转换器、配置校验、唯一性校验及未知动态列策略。
- 单 Sheet 批量请求没有未知动态列选项，沿用忽略未知动态表头的默认行为；Workbook Sheet 请求继续遵循 `FailOnUnknownDynamicColumns`。
- 完整 Workbook Import、动态关系、Entity Layout、Failure Workbook 不在本轮扩展。

## 实施步骤

1. 将批量表头绑定扩展为固定列和动态列绑定，校验动态属性数量、可写性及定义冲突。
2. 在同步、异步行物化中执行动态列原始校验、转换后校验和唯一性跟踪，并将每行动态值写入独立字典。
3. 接入 Workbook 原生校验快照，保留原始工作表、行列坐标和不支持策略。
4. 增加单 Sheet 与 Workbook 的同步/异步测试，覆盖定义匹配、别名、未知列、转换器、校验、唯一性、空值和批次隔离。
5. 更新 Provider 能力说明、执行记录和生产符号到测试方法追溯；不更新公共 API 基线。

## 验收

- 动态列实体可以跨多个批次成功交付，每个实体的字典只包含自身动态值。
- 动态转换器、校验和唯一性失败返回结构化错误，失败行不进入成功实体批次，唯一值预留正确回滚。
- Workbook 批量的内联/名称列表原生校验对动态列保持正确坐标。
- 同步与异步结果一致，取消、回调异常、输入流所有权和暂存清理保持原契约。
- 受影响项目双 TFM 构建与 ExcelDataReader 测试通过，文档测试、`git diff --check`、编码检查通过。
