<!-- AI_EXECUTION_STATUS: COMPLETED -->
AI_TASK_ID: BING-OFFICES-BATCH-WORKBOOK-VALIDATION-20260930-001
AI_EXECUTION_FINISHED_AT: 2026-09-30T11:33:00.000Z
AI_EXECUTION_STARTED_AT: 2026-09-30T11:10:25.161Z

# 执行记录

任务：`BING-OFFICES-BATCH-WORKBOOK-VALIDATION-20260930-001`

状态：已完成。

## 变更

- 为 ExcelDataReader 增加 XLSX OOXML 原生校验快照，支持固定列分批导入中的内联列表、名称列表、整数、小数、日期、时间和文本长度规则。
- 单 Sheet 分批导入使用固定失败策略；Workbook 分批导入遵循请求的 `ExcelUnsupportedFeaturePolicy`，并覆盖同步、异步路径。
- 原生校验在类型转换前执行，错误保留原始工作表、行列坐标；动态列、关系、完整 Import 和 Failure Workbook 仍保持原有边界。
- XLS/XLSB 在需要 Workbook 原生校验时明确返回不支持，不增加 NPOI/ClosedXML 依赖，也未新增公共 API。
- 补充 provider 能力说明、边界说明、职责级测试和生产符号到测试方法的追溯。

## 验证

- `Bing.Offices.ExcelDataReader` Release 构建：net6.0、net8.0 均通过，0 警告、0 错误。
- `Bing.Offices.ExcelDataReader.Tests`：net6.0、net8.0 均通过，52/52，筛选 `Category!=Large`。
- `Bing.Offices.Docs.Tests`：net8.0 通过，19/19，筛选 `Category!=Large`。
- 已覆盖内联/名称列表、数值、日期、文本长度、Workbook 批量异步以及 XLS 预检拒绝场景。
- 已执行 `git diff --check`；本任务改动文件完成 UTF-8、BOM、LF 和末尾换行检查。
